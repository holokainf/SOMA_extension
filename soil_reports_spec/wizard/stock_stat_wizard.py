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
import xlwt
from xlwt import *
from datetime import datetime
import tempfile
import xlsxwriter as xlw
import pandas as pd
import numpy as np
import sys
import calendar as cal


reload(sys)
sys.setdefaultencoding('utf8')



# les exports sont enregistrés dans cette table afin d'être téléchargé.
class stock_stats_wizard_export(orm.TransientModel):
    _name = "stock.stats.wizard.export"
    _description = "Rapport de stock"

    _columns = {
        'name': fields.char('Filename'),
        'data': fields.binary('File'),
    }

stock_stats_wizard_export()

class stock_stats_wizard(orm.TransientModel):
    _name = 'stock.stats.wizard'

    _columns = {
        'date_start': fields.date('Date de départ',required=True),
        'date_end': fields.date('Date de fin',required=True),
        'stats_type': fields.selection([
            ('depot', 'Depot'),
        ], 'Statistiques par : ', select=True, readonly=False, required=True),

    }
    _defaults = {
        'stats_type': 'depot',
        'date_start': datetime.today().strftime("%Y-%m-%d"),
        'date_end':datetime.today().strftime("%Y-%m-%d")
    }

    def print_report_xls(self, cr, uid, ids, context=None):
        record = self.browse(cr,uid,ids[0])
        type = record.stats_type
        date_start = record.date_start
        date_end = record.date_end

        # On créer un buffer
        buffr=StringIO()
        date_day = datetime.now()

        if type == 'depot':
            nom_fichier='Rapport De Stock Du '+ str(date_start)+' Au '+str(date_end)+'.xls'
            self._get_excel_stock_reporting(cr,uid,date_start,date_end,buffr,nom_fichier)

        # On récupère les données dans le buffer
        out=base64.encodestring(buffr.getvalue())
        # On libère la mémoire
        buffr.close()
        # On enregistre dans le DB le fichier.
        if type == 'depot':
            # wizard_id = self.pool.get('stock.stats.wizard.export').create(cr, uid, {'data':out,'name':fichier,'state':'get'}, context=dict(context, active_ids=ids))
            wizard_id = self.pool.get('stock.stats.wizard.export').create(cr, uid, {'data': out,'name':nom_fichier})
            return {
                        'name':"Export Excel",
                        'view_mode': 'form',
                        'view_id':False,
                        'view_type': 'form',
                        'res_model': 'stock.stats.wizard.export',
                        'res_id': wizard_id,
                        'type': 'ir.actions.act_window',
                        'target': 'new',
                        'domain': '[]',
                        # 'context': dict(context, active_ids=ids)
                    }

    def _get_couple_depot_produit(self,cr,uid):
        depot_dict=[]
        cr.execute(""" select id from product_loction_cmp order by location_id""")
        couple_depot_produit_dict = cr.dictfetchall()
        couple_depot_produit = [i['id'] for i in couple_depot_produit_dict]
        return couple_depot_produit

    def _get_sum_stock_initial(self,cr,uid,date_start,date_end,sens,depot_id,product_id):
        id_depot_clt_frs = self.pool.get('stock.location').search(cr, uid, [('usage', '=', 'supplier')])
        annee = datetime.strptime(date_start, "%Y-%m-%d").year
        if sens == 'out':
            cr.execute("""
                select
                    sum(case when location_dest_id = %s then converted_quantity else product_qty end) as total_qtt
                from stock_move
                where date_expected < %s
                    and extract(year from date_expected) = %s
                    and location_id = %s
                    and product_id = %s
            """,(id_depot_clt_frs[0],date_start,annee,depot_id,product_id))
        elif sens == 'in':
            cr.execute("""
                select
                    sum(case when location_id = %s then converted_quantity else product_qty end) as total_qtt
                from stock_move
                where date_expected < %s
                    and extract(year from date_expected) = %s
                    and location_dest_id = %s
                    and product_id = %s
            """,(id_depot_clt_frs[0],date_start,annee,depot_id,product_id))

        qtt_ini_dict = cr.dictfetchall()
        if qtt_ini_dict:
            qtt_ini = [i['total_qtt'] for i in qtt_ini_dict]
            return qtt_ini[0]
        else:
            return

    def _get_mvt(self,cr,uid,operation,sens,date_start,date_end,depot_id,product_id):
        mvt_obj = self.pool.get('stock.move')
        location_obj = self.pool.get('stock.location')
        # print 'operation--------->' + operation + 'sens--------->'+ sens

        if operation == 'inventaire':
            id_depot_ref = self.pool.get('stock.location').search(cr, uid, [('usage', '=', 'inventory')])
        elif  operation == 'achat':
            id_depot_ref = self.pool.get('stock.location').search(cr, uid, [('usage', '=', 'supplier')])
        elif  operation == 'vente':
            id_depot_ref = self.pool.get('stock.location').search(cr, uid, [('usage', '=', 'customer')])
        elif  operation == 'transfert':
            id_depot_ref = self.pool.get('stock.location').search(cr, uid, [('code', '=', 'TMP')])
        elif  operation == 'reclassement':
            id_depot_ref = self.pool.get('stock.location').search(cr, uid, [('is_rec', '=', True)])

        # print 'id_depot_ref: ',location_obj.browse(cr,uid,)

        if id_depot_ref:
            if sens == 'in':
                mvt_ids = mvt_obj.search(cr, uid, [('location_id', '=', id_depot_ref[0]),
                                                   ('location_dest_id', '=', depot_id),
                                                   ('date_expected', '>=', date_start),
                                                   ('date_expected', '<=', date_end),
                                                   ('product_id', '=', product_id),
                                                   ('state', '=', 'done')])
                return mvt_ids
            elif sens == 'out':
                mvt_ids = mvt_obj.search(cr, uid, [('location_id', '=', depot_id),
                                                   ('location_dest_id', '=', id_depot_ref[0]),
                                                   ('date_expected', '>=', date_start),
                                                   ('date_expected', '<=', date_end),
                                                   ('product_id', '=', product_id),
                                                   ('state', '=', 'done')])
                return mvt_ids



    def _get_data_table(self,cr,uid,mvt_ids,sens,operation,product_id,depot_id):
        # print '_get_data_table - mvt_ids :',mvt_ids
        mvt_ids_obj = self.pool.get('stock.move').browse(cr,uid,mvt_ids)
        cmp_obj = self.pool.get('month.stock.qty')
        period_obj = self.pool.get('account.period')

        product_obj = self.pool.get('product.product').browse(cr,uid,product_id)
        product = product_obj.name_template
        depot_obj  = self.pool.get('stock.location').browse(cr,uid,depot_id)
        depot = depot_obj.name

        coeff = 1
        if sens == 'out':
            coeff = -1


        titres = ['ID',
                  'Operation',
                  'Sens',
                  'Dépôt',
                  'Date',
                  'Document',
                  'Article',
                  'Quantité',
                  'Frais Unitaire',
                  'Coût Unitaire Hors Frais',
                  'Coût Unitaire Total',
                  'CMPU Calculé',
                  'N° Ecriture Comptable',
                  'Report Comptable']

        table_liste = []

        for mvt in mvt_ids_obj:

            quantity  = mvt.product_qty
            frais = mvt.frais_transfert
            cout_hors_frais = mvt.price_unit_net

            if operation == 'inventaire':
                document = mvt.name
            elif operation =='achat':
                document = mvt.picking_id.name
                quantity = mvt.converted_quantity
                if mvt.converted_quantity == 0:
                    frais = 0
                    cout_hors_frais = 0
                else:
                    frais = (mvt.frait_achat * mvt.product_qty) / mvt.converted_quantity
                    cout_hors_frais = mvt.price_unit_net * mvt.product_qty / mvt.converted_quantity

            elif operation =='vente':
                document = mvt.picking_id.name
            elif operation == 'transfert':
                document = mvt.internal_picking_id.name
            elif operation == 'reclassement':
                document = mvt.reclassification_id.name


            period_id = period_obj.search(cr, uid,[('name', '=', str(mvt.date_expected[5:7] + "/" + mvt.date_expected[:4]))])


            cmpu_id = cmp_obj.search(cr,uid,[('product_id','=',mvt.product_id.id),('location_id','=',depot_id),('period_id','=',period_id[0])])
            # print 'cmpu_id: ',cmpu_id, ' ** period_id: ',period_id
            if cmpu_id:
                cmpu = cmp_obj.browse(cr, uid, cmpu_id[0]).cmp
            else:
                cmpu = 'N/A'

            ecrt_id = self.pool.get('account.move.line').search(cr, uid, [('stock_move_id', '=', mvt.id), (
            'account_id', '=', mvt.product_id.product_tmpl_id.categ_id.property_stock_valuation_account_id.id)])

            if ecrt_id:
                ecrt = self.pool.get('account.move.line').browse(cr, uid, ecrt_id[0]).debit - self.pool.get(
                    'account.move.line').browse(cr, uid, ecrt_id[0]).credit
                if operation not in 'achat':
                    if mvt.product_qty == 0:
                        ecrt = 0
                    else:
                        ecrt = round(ecrt / mvt.product_qty, 3)
                    ecrt_id = ecrt_id[0]

                else:
                    if mvt.converted_quantity == 0:
                        ecrt = 0
                    else:
                        ecrt = round(ecrt / mvt.converted_quantity, 3)
                    ecrt_id = ecrt_id[0]
            else:
                ecrt = 'Pas de pièce comptable'
                ecrt_id = 0




            table = (mvt.id,
                     operation,
                     sens,
                     depot,
                     mvt.date_expected,
                     document,
                     product,
                     quantity*coeff,
                     frais,
                     cout_hors_frais,
                     frais + cout_hors_frais,
                     cmpu,
                     ecrt_id,
                     ecrt
                     )

            table_liste.append(table)
        # print 'table_liste --->',table_liste
        dt = pd.DataFrame.from_records(table_liste, columns=titres)
        # print df.__len__()
        return dt

    def _get_excel_stock_reporting(self,cr,uid,date_start,date_end,buffr,nom_fichier,context=None):

        operation_list = ['inventaire','achat','vente','transfert','reclassement']
        sens_list = ['out','in']

        obj_couple_depot_produit = self.pool.get('product.loction.cmp')
        ids_couple_depot_produit = obj_couple_depot_produit.browse(cr,uid, self._get_couple_depot_produit(cr,uid))

        df = pd.DataFrame()
        df_list = []
        for i in range(len(ids_couple_depot_produit)):

            id_depot =ids_couple_depot_produit[i].location_id.id
            id_product = ids_couple_depot_produit[i].product_id.id


            for operation in operation_list:
                for sens in sens_list:
                    # print 'operation: '+operation + ' ** Sens: '+sens
                    # print 'Depot: ',str(ids_couple_depot_produit[i].product_id.name_template) + '; Produit: '+ str(ids_couple_depot_produit[i].location_id.name)

                    mvt_ids = self._get_mvt(cr,uid,operation,sens,date_start,date_end,id_depot,id_product)

                    if mvt_ids:
                        df_list.append(self._get_data_table(cr, uid, mvt_ids,sens,operation,id_product,id_depot))

        df = pd.concat(df_list)
        # df['Montant Total']
        writer = pd.ExcelWriter(buffr, engine='xlsxwriter')
        wb = writer.book

        ws1_titre = [{'header':str(x)} for x in df.columns.tolist()]
        ws1 = wb.add_worksheet('Liste des mouvement')
        ws1.add_table(1, 0, len(df), len(df.axes[1]),{'name': 'stock_move',
                                                      'data': df.values,
                                                      'columns':ws1_titre,
                                                      'total_row': True,
                                                      'style': 'Table Style Medium 2'
                                                      }
                      )

        df_group_produit = pd.pivot_table(df,
                       values =['Quantité'],
                       index=['Dépôt','Article','Operation'],
                       aggfunc=np.sum)

        dgp = df_group_produit.unstack(level=2)

        dgp.to_excel(writer, sheet_name='Statistique par produit')

stock_stats_wizard()