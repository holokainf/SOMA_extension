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
# import matplotlib.pyplot as plt
import tempfile

# les exports sont enregistrés dans cette table afin d'être téléchargé.
class stock_stats_wizard_export(orm.TransientModel):
    _name = "stock.stats.wizard.export"
    _description = "Rapport de stock"

    _columns = {
        'name': fields.char('Filename'),
        'data': fields.binary('File'),
        # 'state': fields.selection( [ ('choose','choose'),('get','get')],'state', readonly=True),
    }
    _defaults = {
        # 'state':'choose',
    }
stock_stats_wizard_export()

class stock_stats_wizard(orm.TransientModel):
    _name = 'stock.stats.wizard'

    _columns = {
        'fiscalyear_id': fields.many2one('account.fiscalyear', 'Exercice', required=True),
        'period_from_id': fields.many2one('account.period', 'Période'),
        'date_start': fields.date('Date de départ',required=True),
        'date_end': fields.date('Date de fin',required=True),
        # 'export_canal_culumns': fields.boolean('Exporter les colonnes (Canal parent et canal)'),
        'stats_type': fields.selection([
            ('depot', 'Depot'),
            # ('relation', 'Relation'),
            # ('produit', 'Produit'),
        ], 'Statistiques par : ', select=True, readonly=False, required=True),

    }
    _defaults = {
        'stats_type': 'depot',
        # 'export_canal_culumns': True,
        'fiscalyear_id': lambda self, cr, uid, context:
        self.pool.get('account.fiscalyear').browse(cr, uid, self.pool.get('account.fiscalyear').search(cr, uid, []))[
            -1].id,
        'date_start': datetime(datetime.today().year,datetime.today().month,1).strftime("%Y-%m-%d"),
        'date_end':datetime.today().strftime("%Y-%m-%d"),
    }

    def print_report_xls(self, cr, uid, ids, context=None):
        record = self.browse(cr,uid,ids[0])
        type = record.stats_type
        date_start = record.date_start
        date_end = record.date_end

        # On créer un buffer
        buf=StringIO()
        date_day = datetime.now()

        if type == 'depot':
            self.export_stat_stock(cr, uid, ids,date_start,date_end,buf,context)
            fichier="rapport_stock_depot.xls"
        elif type == 'relation':
            self.export_relation_stock(cr, uid, ids,date_start,date_end,context)
            # fichier="rapport_stock_depot_du_"+str(date_day)+".xls"

        # On récupère les données dans le buffer
        out=base64.encodestring(buf.getvalue())
        # On libère la mémoire
        buf.close()
        # On enregistre dans le DB le fichier.
        if type == 'depot':
            # wizard_id = self.pool.get('stock.stats.wizard.export').create(cr, uid, {'data':out,'name':fichier,'state':'get'}, context=dict(context, active_ids=ids))
            wizard_id = self.pool.get('stock.stats.wizard.export').create(cr, uid, {'data': out,'name':fichier})
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

    def export_stat_stock(self,cr,uid,ids,date_start,date_end,fp, context=None):
        wb = xlwt.Workbook()
        ws = wb.add_sheet('Rapport De Stock',cell_overwrite_ok=True)
        ws1 = wb.add_sheet('Mouvements De Stock', cell_overwrite_ok=True)
        ws2 = wb.add_sheet('Historique des CMPU', cell_overwrite_ok=True)
        self.excel_write_his_cmpu(cr, uid, ws2)

        #######################  Création des styles excel #######################
        font0 = xlwt.Font()
        font0.name = 'Arial'
        font0.colour_index = 2
        font0.bold = True
        font0.height = 350

        font1 = xlwt.Font()
        font1.name = 'Arial'
        font1.colour_index = 0
        font1.bold = True

        font6 = xlwt.Font()
        font6.name = 'Arial'
        font6.colour_index = 0
        font6.bold = True

        font2 = xlwt.Font()
        font2.name = 'Arial'
        font2.colour_index = 2
        font2.bold = True
        font2.height = 280

        font3 = xlwt.Font()
        font3.name = 'Arial'
        font3.colour_index = 2
        font3.bold = True
        font3.height = 200

        font4 = xlwt.Font()
        font4.name = 'Arial'
        font4.colour_index = xlwt.Style.colour_map['white']
        font4.bold = True

        font5 = xlwt.Font()
        font5.name = 'Arial'
        font0.colour_index = 2

        style0 = xlwt.XFStyle()
        style0.font = font0

        style1 = xlwt.XFStyle()
        style1.font = font2

        alignement = xlwt.Alignment()
        alignement.horz = xlwt.Alignment.HORZ_RIGHT

        alignement2 = xlwt.Alignment()
        alignement2.horz = xlwt.Alignment.HORZ_CENTER

        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = xlwt.Style.colour_map['blue']

        pattern_grey = xlwt.Pattern()
        pattern_grey.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern_grey.pattern_fore_colour = xlwt.Style.colour_map['blue_gray']

        pattern_blue= xlwt.Pattern()
        pattern_blue.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern_blue.pattern_fore_colour = xlwt.Style.colour_map['blue']

        style2 = xlwt.XFStyle()
        style2.num_format_str = '#,##0.00'
        style2.font = font1
        style2.alignment = alignement
        style2.pattern = pattern

        borders = Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1

        style3 = xlwt.XFStyle()
        style3.font = font4
        style3.borders = borders

        style4 = xlwt.XFStyle()
        style4.font = font5
        style4.alignment = alignement

        style5 = xlwt.XFStyle()
        style5.font = font5
        style5.num_format_str = '#,##0.00'

        style6 = xlwt.XFStyle()
        style6.font = font3

        style7 = xlwt.XFStyle()
        style7.font = font1
        style7.borders = borders
        style7.alignment = alignement2

        style8 = xlwt.XFStyle()
        style8.font = font1

        style9 = xlwt.XFStyle()
        style9.font = font5

        font10 = xlwt.Font()
        font10.name = 'Arial'
        font10.colour_index = 63
        font10.bold = False

        style10 = xlwt.XFStyle()
        style10.num_format_str = '#,##0.00'
        style10.font = font10
        style10.alignment = alignement

        style11 = xlwt.XFStyle()
        style11.num_format_str = '#,##0.00'
        style11.font = font6
        style11.font.colour_index = xlwt.Style.colour_map['white']
        style11.alignment = alignement
        style11.pattern = pattern_grey

        style12 = xlwt.XFStyle()
        style12.font = font6
        style12.num_format_str = '#,##0.00'
        style12.font.colour_index = xlwt.Style.colour_map['white']
        style12.alignment = alignement
        style12.pattern = pattern_blue

        style13 = xlwt.XFStyle()
        style13.font = font6
        style13.font.colour_index = xlwt.Style.colour_map['white']
        style10.num_format_str = '#,##0.00'
        style13.pattern = pattern_grey

        style14 = xlwt.XFStyle()
        style14.font = font6
        style10.num_format_str = '#,##0.00'
        style14.font.colour_index = xlwt.Style.colour_map['white']
        style14.pattern = pattern_blue


        style_titre1 = xlwt.easyxf('font: height 400, bold 1, color black, italic 0, underline single;'
                                'alignment: horizontal center, vertical center;'
                                ' borders: left thin, right thin, top thick, ''bottom thick, top_color green;'
                                'pattern: pattern solid, fore_color yellow', num_format_str = 'YYYY-MM-DD')


        #########################
        header_p = [unicode('Dépôt', "utf8"),
                    unicode('Article', "utf8"),
                    unicode('Stock au '+str(date_start), "utf8"),
                    unicode('Inventaire', "utf8"),
                    unicode('Achat', "utf8"),
                    unicode('Vente', "utf8"),
                    unicode('Transfert', "utf8"),
                    unicode('Reclassement', "utf8"),
                    unicode('Stock au '+str(date_end), "utf8"),
                    ]
        header_mvt_stk = [unicode('ID', "utf8"),unicode('Opération', "utf8"),unicode('Sens', "utf8"),
                          unicode('Dépôt', "utf8"),unicode('Date', "utf8"),
                          unicode('Document', "utf8"),unicode('Article', "utf8"),
                          unicode('Quantité', "utf8"),unicode('Frais Unitaire', "utf8"),
                          unicode('Coût Uniaire Hors Frais', "utf8"),unicode('Coût Uniaire Total', "utf8"),unicode('CMPU Calculé', "utf8"),
                          unicode("N° Ecriture Comptable",'utf8'),unicode("Report Comptable",'utf8'),]

        operation_list = ['inventaire','achat','vente','transfert','reclassement']
        sens_list = ['out','in']


        ws_titre = {'row':5,'col':0}
        ws_data = {'row':6,'col':2}
        ws1_titre = {'row':0,'col':0}
        ws1_data = {'row':1,'col':2}
        ligne = 1


        ws.write_merge(0, 0, 0,len(header_p)-1, 'Rapport de Stock Par Depot du '+str(date_start)+ ' Au ' + str(date_end),style_titre1)

        # Ecriture des titres
        for i in range(len(header_p)):
            ws.col(i).width = (len(header_p[i])*500)
            ws.write(ws_titre.get('row'),i,header_p[i], style7)
        for i in range(len(header_mvt_stk)):
            ws1.col(i).width = (len(header_mvt_stk[i])*500)
            ws1.write(0,i,header_mvt_stk[i], style7)

        obj_couple_depot_produit = self.pool.get('product.loction.cmp')
        ids_couple_depot_produit = obj_couple_depot_produit.browse(cr,uid, self._get_couple_depot_produit(cr,uid))

        for i in range(len(ids_couple_depot_produit)):

            id_depot =ids_couple_depot_produit[i].location_id.id
            id_product = ids_couple_depot_produit[i].product_id.id

            ws.write(ws_data.get('row')+i,0,ids_couple_depot_produit[i].location_id.name)
            ws.write(ws_data.get('row')+i, 1, ids_couple_depot_produit[i].product_id.name)

            #########  STOCK INITIAL ########
            qty_initial = 0
            for sens in sens_list:
                qty = self._get_sum_stock_initial(cr, uid, date_start, date_end, sens, id_depot, id_product)
                if qty:
                    if sens == 'in':
                        qty_initial += qty
                    elif sens == 'out':
                        qty_initial -=qty
            ws.write(ws_data.get('row') + i, ws_data.get('col'), qty_initial)

            for operation in operation_list:
                qty = 0
                for sens in sens_list:
                    mvt_ids = self._get_mvt(cr,uid,operation,sens,date_start,date_end,id_depot,id_product)
                    if mvt_ids:
                        ligne = self.excel_write_mvt(cr, uid, ws1, mvt_ids, operation, sens, ligne)
                        # print 'ligne = ',ligne
                        qty_tmp = self._get_sum_mvt(cr, uid, mvt_ids, operation)
                        if mvt_ids and sens == 'in':
                            qty += qty_tmp
                        if mvt_ids and sens == 'out':
                            qty -= qty_tmp
                ws.write(ws_data.get('row') + i,ws_data.get('col') + operation_list.index(operation)+1,qty)
            #########  Stock_final  ########
            ws.write(ws_data.get('row') + i, 8, xlwt.Formula("SUM(C%d:H%d)" % (ws_data.get('row')+i+1,ws_data.get('row')+i+1)))

        wb.save(fp)

    def _get_couple_depot_produit(self,cr,uid):
        depot_dict=[]
        cr.execute(""" select id from product_loction_cmp order by location_id""")
        couple_depot_produit_dict = cr.dictfetchall()
        couple_depot_produit = [i['id'] for i in couple_depot_produit_dict]
        # print 'couple_depot_produit ----->',couple_depot_produit_dict
        return couple_depot_produit

    def _get_product(self,cr,uid):
        cr.execute(""" select distinct product_id from product_loction_cmp """)
        product_dict = cr.dictfetchall()
        product = [i['product_id'] for i in product_dict]
        return product

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

    # def _get_sum_mvt(self,cr,uid,ids_mvt,mvt_type):
    #     sum_qty = 0
    #     if ids_mvt:
    #         if mvt_type not in ['achat']:
    #             cr.execute("""
    #                 select sum(product_qty) as quantity from stock_move where id in %s
    #             """,(tuple(ids_mvt,),))
    #         elif mvt_type == 'achat':
    #             cr.execute("""
    #                 select sum(converted_quantity) as quantity from stock_move where id in %s
    #             """,(tuple(ids_mvt,),))
    #         sum_qty =  cr.dictfetchall()[0].values()[0]
    #     return sum_qty

    def _get_mvt(self,cr,uid,operation,sens,date_start,date_end,depot_id,product_id):
        mvt_obj = self.pool.get('stock.move')
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

        if not id_depot_ref:
            raise osv.except_osv(_('Error!'), _('Aucune location trouver pour cette operation'))
        else:
            if sens == 'in':
                mvt_ids = mvt_obj.search(cr, uid, [('location_id', '=', id_depot_ref[0]),
                                                   ('location_dest_id', '=', depot_id),
                                                   ('date_expected', '>=', date_start),
                                                   ('date_expected', '<=', date_end),
                                                   ('product_id', '=', product_id),
                                                   ('state', '=', 'done')])
            elif sens == 'out':
                mvt_ids = mvt_obj.search(cr, uid, [('location_id', '=', depot_id),
                                                   ('location_dest_id', '=', id_depot_ref[0]),
                                                   ('date_expected', '>=', date_start),
                                                   ('date_expected', '<=', date_end),
                                                   ('product_id', '=', product_id),
                                                   ('state', '=', 'done')])

        return mvt_ids

    def _get_data_table(self,cr,uid,mvt_ids,sens,operation,product_id,depot_id):

        mvt_ids_obj = self.pool.get('stock_move').browse(cr,uid,mvt_ids)
        cmp_obj = self.pool.get('month.stock.qty')
        period_obj = self.pool.get('account.period')

        product_obj = self.pool.get('product.product').browse(cr,uid,product_id)
        product = product_obj.name_template
        depot_obj  = self.pool.get('stock.location').browse(cr,uid,depot_id)
        depot = depot_obj.name

        coeff = 1
        if sens = 'out':
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
            if operation == 'inventaire':
                document = mvt.name
            elif operation ='achat':
                document = mvt.picking_id.name
                quantity = mvt.converted_quantity
                frais = (mvt.frait_achat * mvt.product_qty) / mvt.converted_quantity
                cout_hors_frais = mvt.price_unit_net * mvt.product_qty / mvt.converted_quantity
            elif operation ='vente':
                document = mvt.picking_id.name
            elif operation == 'transfert':
                document = mvt.internal_picking_id.name
            elif operation == 'reclassement':
                document = mvt.reclassification_id.name


            period_id = period_obj.search(cr, uid,
                                          [('name', '=', str(mvt.date_expected[5:7] + "/" + mvt.date_expected[:4]))])


            cmpu_id = cmp_obj.search(cr,uid,[('product_id','=',mvt.product_id.id),('location_id','=',depot),('period_id','=',period_id[0])])
            cmpu = cmp_obj.browse(cr, uid, cmpu_id[0])

            ecrt_id = self.pool.get('account.move.line').search(cr, uid, [('stock_move_id', '=', mvt.id), (
            'account_id', '=', mvt.product_id.product_tmpl_id.categ_id.property_stock_valuation_account_id.id)])

            if ecrt_id:
                ecrt = self.pool.get('account.move.line').browse(cr, uid, ecrt_id[0]).debit - self.pool.get(
                    'account.move.line').browse(cr, uid, ecrt_id[0]).credit
                if operation not in 'achat':
                    ecrt = round(ecrt / mvt.product_qty, 3)
                else:
                    ecrt = round(ecrt / mvt.converted_quantity, 3)


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
                     ecrt_id[0],
                     ecrt
                     )

        table_liste.append(table)

        return pd.DataFrame.from_records(table_liste, columns=titres)


    # def excel_write_mvt(self,cr,uid,ws1,mvt_ids,operation,sens,ligne):
    #     if mvt_ids:
    #         mvt_obj = self.pool.get('stock.move')
    #         titre = ['id','operation','sens','depot','date','document','article','quantite','frais','cout_hors_frais','cout_total','cmpu','Ligne Comptable','Report comptable']
    #         cmp_obj = self.pool.get('month.stock.qty')
    #         period_obj = self.pool.get('account.period')
    #
    #         for mvt in mvt_obj.browse(cr, uid, mvt_ids):
    #             period_id =  period_obj.search(cr,uid,[('name','=',str(mvt.date_expected[5:7]+"/"+mvt.date_expected[:4]))])
    #             # print 'period_id: ',period_id[0]
    #             if sens == 'in':
    #                 depot = mvt.location_id.id
    #                 tt = mvt.location_id.name
    #             else:
    #                 depot = mvt.location_id.id
    #                 tt = mvt.location_id.name
    #             # print 'product_id:',mvt.product_id.name_template
    #             # print 'depot_id: ' + tt + "sens: "+ sens
    #
    #             cmpu_id = cmp_obj.search(cr,uid,[('product_id','=',mvt.product_id.id),('location_id','=',depot),('period_id','=',period_id[0])])
    #             # print 'ligne inventaire in ------> ', cmpu_id
    #
    #             qty = 0
    #             frais = 0
    #             cout_hors_frais = 0
    #             for t in titre:
    #
    #                 if t == 'id':ws1.write(ligne, titre.index(t), mvt.id)
    #                 if t == 'operation':ws1.write(ligne, titre.index(t), operation)
    #                 if t == 'sens': ws1.write(ligne, titre.index(t), sens)
    #                 if t == 'depot':
    #                     if sens == 'in':ws1.write(ligne, titre.index(t), mvt.location_dest_id.name)
    #                     elif sens == 'out':ws1.write(ligne, titre.index(t), mvt.location_id.name)
    #                 if t == 'date':ws1.write(ligne, titre.index(t), mvt.date_expected)
    #                 if t == 'document':
    #                     if operation == 'inventaire':ws1.write(ligne, titre.index(t), mvt.name)
    #                     elif operation in ['achat', 'vente']:ws1.write(ligne, titre.index(t), mvt.picking_id.name)
    #                     elif operation == 'transfert':ws1.write(ligne, titre.index(t), mvt.internal_picking_id.name)
    #                     elif operation == 'reclassement':ws1.write(ligne, titre.index(t), mvt.reclassification_id.name)
    #                 if t == 'article':ws1.write(ligne, titre.index(t), mvt.product_id.name_template)
    #                 if t == 'quantite':
    #                     if sens == 'in':
    #                         qty+=mvt.product_qty
    #                     elif sens == 'out':
    #                         qty += mvt.product_qty * (-1)
    #                     ws1.write(ligne, titre.index(t),qty)
    #
    #                 if t == 'frais':
    #                     if operation not in 'achat':
    #                         frais += mvt.frais_transfert
    #                     elif operation in ['achat']:
    #                         frais += (mvt.frait_achat * mvt.product_qty) / mvt.converted_quantity
    #                     ws1.write(ligne, titre.index(t),round(frais,3))
    #
    #                 if t == 'cout_hors_frais':
    #                     if operation not in 'achat':
    #                         cout_hors_frais =  mvt.price_unit_net
    #                     elif operation in ['achat']:
    #                         if mvt.converted_quantity:
    #                             cout_hors_frais =  mvt.price_unit_net * mvt.product_qty / mvt.converted_quantity
    #                         else:
    #                             cout_hors_frais = 0
    #                     ws1.write(ligne, titre.index(t), round(cout_hors_frais,3))
    #                 if t == 'cout_total':
    #                     ws1.write(ligne, titre.index(t), round(frais + cout_hors_frais,3) )
    #
    #                 if t =='cmpu' and sens == 'out':
    #                     if cmpu_id:
    #                         cmpu = cmp_obj.browse(cr, uid, cmpu_id[0])
    #                         ws1.write(ligne, titre.index(t), round(cmpu.cmp,3))
    #                     else:
    #                         ws1.write(ligne, titre.index(t), 'non evalue')
    #                 if t == 'Report comptable':
    #                     ecrt_id =self.pool.get('account.move.line').search(cr,uid,[('stock_move_id','=',mvt.id),('account_id','=',mvt.product_id.product_tmpl_id.categ_id.property_stock_valuation_account_id.id)])
    #                     # print 'ecrt_id  ----->',ecrt_id
    #                     if ecrt_id:
    #                         ecrt = self.pool.get('account.move.line').browse(cr,uid,ecrt_id[0]).debit - self.pool.get('account.move.line').browse(cr,uid,ecrt_id[0]).credit
    #                         if operation not in 'achat':
    #                             ws1.write(ligne, titre.index(t), round(ecrt/mvt.product_qty,3))
    #                         else:
    #                             ws1.write(ligne, titre.index(t), round(ecrt/mvt.converted_quantity,3))
    #                         ws1.write(ligne, titre.index(t)-1, ecrt_id[0])
    #                     else:
    #                         ws1.write(ligne, titre.index(t), 'Pas de piece comptable')
    #                 # if t == 'Ligne Comptable':
    #                 #     ws1.write(ligne, titre.index(t), ecrt_id[0])
    #             ligne += 1
    #
    #         return ligne
    #
    # def excel_write_his_cmpu(self,cr,uid,ws2):
    #     header_hist_cmp = [unicode('Periode', "utf8"), unicode('Dépôt', "utf8"),
    #                       unicode('Produit', "utf8"), unicode('Stock De Fin de Mois', "utf8"),
    #                       unicode('Coût unitaire', "utf8"), unicode('Valeur Totale Des Stocks', "utf8")]
    #
    #     cmp_obj = self.pool.get('month.stock.qty')
    #     cr.execute(""" select id from month_stock_qty order by period_id,product_id""")
    #     cmp_dict = cr.dictfetchall()
    #     cmp_ids = [i['id'] for i in cmp_dict]
    #
    #
    #     ws2_titre = {'row':0,'col':0}
    #     ws2_data = {'row':1,'col':0}
    #
    #     for i in range(len(header_hist_cmp)):
    #         ws2.col(i).width = (len(header_hist_cmp[i])*500)
    #         ws2.write(ws2_titre.get('row'),ws2_titre.get('col')+i,header_hist_cmp[i])
    #
    #     i =0
    #     for cmp in cmp_obj.browse(cr,uid,cmp_ids):
    #         ws2.write(ws2_data.get('row') + i,0,cmp.period_id.date_stop)
    #         ws2.write(ws2_data.get('row')+i, 1, cmp.location_id.name)
    #         ws2.write(ws2_data.get('row')+i, 2, cmp.product_id.name_template)
    #         ws2.write(ws2_data.get('row')+i, 3, cmp.end_qty)
    #         ws2.write(ws2_data.get('row')+i, 4, cmp.cmp)
    #         ws2.write(ws2_data.get('row')+i, 5, cmp.cmp * cmp.end_qty)
    #         i+=1
    #
    # # def _get_cmpu(self,cr,uid,depot_id,product_id,mois):
    # #     cmp_obj = self.pool.get('month.stock.qty')
    # #     cr.execute(""" select id from month_stock_qty order by period_id,product_id""")
    # #     cmp_dict = cr.dictfetchall()
    # #     cmp_ids = [i['id'] for i in cmp_dict]
    #
    # # def cmpu_hist_pdf(self,cr,uid):
    # #
    # #     x = [1, 2, 3, 4]
    # #     y = [4, 7, 9, 8]
    # #     plt.plot(x, y)
    # #     pic_file_name = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
    # #     plt.savefig('/home/user/pic.png')
    # #     pic_data = open('/home/user/pic.png', 'rb').read()
    # #     self.write({'Pic': base64.encodestring(pic_data)})
    #
    # # def export_relation_stock(self,cr, uid, ids,date_start,date_end,context):
    # #     operation_list = ['inventaire','achat','vente','transfert','reclassement']
    # #     sens_list = ['out','in']
    # #     obj_couple_depot_produit = self.pool.get('product.loction.cmp')
    # #     ids_couple_depot_produit = obj_couple_depot_produit.browse(cr,uid, self._get_couple_depot_produit(cr,uid))
    # #
    # #     for i in range(len(ids_couple_depot_produit)):
    # #
    # #         id_depot =ids_couple_depot_produit[i].location_id.id
    # #         id_product = ids_couple_depot_produit[i].product_id.id
    # #
    # #         for operation in operation_list:
    # #             for sens in sens_list:
    # #                 mvt_ids = self._get_mvt(cr,uid,operation,sens,date_start,date_end,id_depot,id_product)
    # #                 self.trajet(cr, uid, mvt_ids, operation, sens)
    # #                 # if mvt_ids:
    # #                 #
    # #                 #     if mvt_ids and sens == 'in':
    # #                 #         # qty += qty_tmp
    # #                 #     if mvt_ids and sens == 'out':
    # #                 #         # qty -= qty_tmp
    # #
    # #
    # # def trajet(self,cr,uid,mvt_ids,operation,sens):
    # #     mvt_obj = self.pool.get('stock.move').browse(cr,uid,mvt_ids)
    # #     lst_mvt=set()
    # #     for mvt in mvt_obj:
    # #         trajet = []
    # #         if operation == 'transfert':
    # #            trajet =(mvt.product_id.id,
    # #                     mvt.internal_picking_id.location_src_id.id,
    # #                     mvt.internal_picking_id.location_dest_id.id,
    # #                     mvt.product_id.id
    # #                     )
    # #         if trajet:
    # #             lst_mvt.add(trajet)
    # #             print trajet
    # #     if lst_mvt:
    # #         print lst_mvt
    # #             # print trajet
    # #     # mvt_set = set(lst_mvt)
stock_stats_wizard()