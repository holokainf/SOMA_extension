
# -*- coding: utf-8 -*-
##############################################################################
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) 2004-2010 Tiny SPRL (http://tiny.be). All Rights Reserved
#
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see http://www.gnu.org/licenses/.
#
######################################################################

from openerp.osv import fields, orm, osv
from tools.translate import _
from cStringIO import StringIO
import base64
from datetime import datetime
# import matplotlib.pyplot as plt
import tempfile
import xlsxwriter as xlw
import pandas as pd
import numpy as np

# les exports sont enregistrés dans cette table afin d'être téléchargé.
class invoice_report_wizard_export(orm.TransientModel):
    _name = "invoice.report.wizard.export"
    _description = "Statistique des ventes"

    _columns = {
        'name': fields.char('Filename'),
        'data': fields.binary('File')
    }

invoice_report_wizard_export()

class invoice_report_wizard(orm.TransientModel):
    _name = 'invoice.report.wizard'

    _columns = {
        # 'fiscalyear_id': fields.many2one('account.fiscalyear', 'Exercice', required=True),
        # 'period_from_id': fields.many2one('account.period', 'Période'),
        'date_start': fields.date('Date de départ',required=True),
        'date_end': fields.date('Date de fin',required=True),
        # 'export_canal_culumns': fields.boolean('Exporter les colonnes (Canal parent et canal)'),
        # 'stats_type': fields.selection([
        #     ('depot', 'Depot'),
        #     # ('relation', 'Relation'),
        #     # ('produit', 'Produit'),
        # ], 'Statistiques par : ', select=True, readonly=False, required=True),

    }
    _defaults = {
        # 'stats_type': 'depot',
        # 'export_canal_culumns': True,
        # 'fiscalyear_id': lambda self, cr, uid, context:
        # self.pool.get('account.fiscalyear').browse(cr, uid, self.pool.get('account.fiscalyear').search(cr, uid, []))[
        #     -1].id,
        'date_start': datetime(datetime.today().year,datetime.today().month,1).strftime("%Y-%m-%d"),
        'date_end':datetime.today().strftime("%Y-%m-%d"),
    }

    def print_report_xls(self, cr, uid, ids, context=None):
        record = self.browse(cr,uid,ids[0])
        date_start = record.date_start
        date_end = record.date_end

        # On créer un buffer
        buffr=StringIO()

        # self.export_stat_stock(cr, uid, ids,date_start,date_end,buf,context)
        nom_fichier = "Statistique des ventes du " + str(date_start) + " au " + str(date_end) + ".xlsx"
        self.export_stat_invoice(cr, uid, ids,date_start,date_end,buffr,nom_fichier,context)

        # print 'kkkkkkkkkkkkkkkk'

        # On récupère les données dans le buffer
        out=base64.encodestring(buffr.getvalue())
        # print 'sssssssss: ',out
        # On libère la mémoire
        buffr.close()
        # On enregistre dans le DB le fichier.
        wizard_id = self.pool.get('invoice.report.wizard.export').create(cr, uid, {'data': out,'name':nom_fichier})
        return {
            'name': "Export Excel",
            'view_mode': 'form',
            'view_id': False,
            'view_type': 'form',
            'res_model': 'invoice.report.wizard.export',
            'res_id':  wizard_id,
            'type': 'ir.actions.act_window',
            'target': 'new','domain': '[]'
            }

    def _get_invoice_ids(self,cr,uid,date_start,date_end):
        inv_obj = self.pool.get('account.invoice')
        inv_ids = inv_obj.search(cr,uid,[('type','like','out_'),
                                         ('date_invoice','>=',date_start),
                                         ('date_invoice','<=',date_end),
                                         ('state','in',('open','paid'))
                                         ])
        return  inv_ids

    def _get_invoice_ids_line(self, cr, uid, inv_ids):
        inv_line_obj = self.pool.get('account.invoice.line')
        inv_line_ids = inv_line_obj.search(cr,uid,[('invoice_id','in',inv_ids)])
        return inv_line_ids

    def _get_currency(self,cr,uid,pc_id):
        ecrt_model = self.pool.get('account.move.line')
        erct_ids = ecrt_model.search(cr,uid,[('move_id','=',pc_id)])
        if erct_ids:
            ecrt = ecrt_model.browse(cr,uid,erct_ids[0])
            return abs(ecrt.debit - ecrt.credit)/ abs(ecrt.amount_currency)

    def _get_data_table(self,cr,uid,inv_line_ids):
        inv_line_obj = self.pool.get('account.invoice.line').browse(cr,uid,inv_line_ids)

        devise_ids = self.pool.get('res.currency').search(cr, uid, [('base', '=', True)])
        devise_defaut = devise_ids[0]

        titres = ['Type',
                  'Date',
                  'Facture',
                  'Origine',
                  'Client',
                  'Produit',
                  'Quantite',
                  'Unite',
                  'Densite',
                  'Poids en TM',
                  'Prix unitaire',
                  'Total Hors Taxe En Devise',
                  'Devise',
                  'Taux de change',
                  'Total Hors Taxe En Monnaie Locale']


        table_liste = []
        for line in inv_line_obj:
            table = ()

            if line.invoice_id.type == 'out_refund':
                coeff = -1
            else:
                coeff = 1

            if line.invoice_id.currency_id.id != devise_defaut:
                currency_rate = self._get_currency(cr, uid, line.invoice_id.move_id.id)
            else:
                currency_rate = 1

            if line.invoice_id.type == 'out_invoice':
                inv_type = 'Facture'
            elif line.invoice_id.type == 'out_refund':
                inv_type = 'Avoir'

            table = (inv_type,
                    line.invoice_id.date_invoice,
                    line.invoice_id.number,
                    line.invoice_id.name,
                    line.invoice_id.partner_id.name,
                    line.product_id.name_template,
                    line.quantity * coeff,
                    line.uos_id.name,
                    line.densite,
                    line.quantity * coeff * line.densite,
                    line.price_unit,
                    line.price_subtotal * coeff,
                    line.invoice_id.currency_id.name,
                    currency_rate,
                    currency_rate *line.price_subtotal * coeff)
            table_liste.append(table)
            # print table_liste
        return pd.DataFrame.from_records(table_liste, columns=titres)

    def export_stat_invoice(self,cr,uid,ids,date_start,date_end,buffr,nom_fichier,context=None):
        # wb = xlw.Workbook(buffr)
        # wb.set_properties({
        #     'title': 'Reporting Des Ventes',
        #     'author': 'Aly Kane',
        #     'company': 'STAR OIL GROUP',
        #     'comments': 'Created with Python and XlsxWriter'})
        # ws1 = wb.add_worksheet('Liste des Factures')

        inv_ids = self._get_invoice_ids(cr, uid, date_start, date_end)
        inv_line_ids = self._get_invoice_ids_line(cr,uid,inv_ids)

        df = self._get_data_table(cr, uid, inv_line_ids)
        df_group_produit = df.groupby(['Produit'])['Total Hors Taxe En Monnaie Locale','Quantite'].agg([np.sum])
        df_group_client = df.groupby(['Client'])['Total Hors Taxe En Monnaie Locale','Quantite'].agg([np.sum])

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(buffr, engine='xlsxwriter')

        # Write the data frame to the BytesIO object.
        df.to_excel(writer, index=False, sheet_name='Liste des Factures')
        df_group_produit.to_excel(writer,sheet_name='Statistique par produit')
        df_group_client.to_excel(writer,sheet_name='Statistique par Client')

        workbook = writer.book
        # worksheet = writer.sheets['Sheet1']

        # Add some cell formats.
        number_format = workbook.add_format({'num_format': '# ##0'})
        money_format = workbook.add_format({'num_format': '# ##0 "CFA"'})

        writer.save()





invoice_report_wizard()


# def _write_invoice_report_wizard(self, cr, uid, inv_line_ids, ws1, wb):
#     inv_line_obj = self.pool.get('account.invoice.line').browse(cr, uid, inv_line_ids)
#
#     merge_format = wb.add_format({'align': 'center', 'bold': 1,
#                                   'underline': 1,
#                                   'bg_color': '#3d5291',
#                                   'font_size': 20,
#                                   'font_color': 'white',
#                                   })
#
#     number_format = wb.add_format({'num_format': '# ##0'
#                                    })
#
#     money_format = wb.add_format({'num_format': '# ##0 "XOF"'
#                                   })
#
#     formula = '=([@[Taux de change]]*[@[Total Hors Taxe En Devise]])'
#
#     titres = [{'header': 'Type'},
#               {'header': 'Date'},
#               {'header': 'Facture'},
#               {'header': 'Origine'},
#               {'header': 'Client'},
#               {'header': 'Produit'},
#               {'header': 'Quantite', 'total_function': 'sum', 'format': number_format},
#               {'header': 'Prix unitaire'},
#               {'header': 'Total Hors Taxe En Devise', 'format': number_format},
#               {'header': 'Devise'},
#               {'header': 'Taux de change'},
#               {'header': 'Total Hors Taxe En Monnaie Locale', 'formula': formula, 'total_function': 'sum',
#                'format': money_format}]
#
#     devise_ids = self.pool.get('res.currency').search(cr, uid, [('base', '=', True)])
#     devise_defaut = devise_ids[0]
#
#     table_liste = []
#     for line in inv_line_obj:
#         table = []
#
#         if line.invoice_id.type == 'out_refund':
#             coeff = -1
#         else:
#             coeff = 1
#
#         if line.invoice_id.currency_id.id != devise_defaut:
#             currency_rate = self._get_currency(cr, uid, line.invoice_id.move_id.id)
#         else:
#             currency_rate = 1
#
#         if line.invoice_id.type == 'out_invoice':
#             inv_type = 'Facture'
#         elif line.invoice_id.type == 'out_refund':
#             inv_type = 'Avoir'
#
#         table.append(inv_type)
#         table.append(line.invoice_id.date_invoice)
#         table.append(line.invoice_id.number)
#         table.append(line.invoice_id.name)
#         table.append(line.invoice_id.partner_id.name)
#         table.append(line.product_id.name_template)
#         table.append(line.quantity * coeff)
#         table.append(line.price_unit)
#         table.append(line.price_subtotal * coeff)
#         table.append(line.invoice_id.currency_id.name)
#         table.append(currency_rate)
#         table_liste.append(table)
#
#     # currency_format = wb.add_format({'num_format': 'CFA #,##0'})
#     ws1.merge_range(0, 0, 0, len(table), 'Liste des factures', merge_format)
#     ws1.add_table(1, 0, len(table_liste), len(table),
#                   # {'name': 'factures', 'data': table_liste, 'columns': titres, 'total_row': True})