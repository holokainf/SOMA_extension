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
from openerp.osv import fields, orm
from osv import fields
from tools.translate import _
from datetime import datetime ,timedelta , date
from operator import itemgetter
import numpy as np
import pandas as pd


class cmpu_compute(orm.TransientModel):

    # def _get_peiriod(self,cr,uid):

    _name = 'cmpu.compute'

    _columns = {
        'period_id': fields.many2one('account.period', "Period", required=True,
            help="Period for which recalculation will be done."),
        'location_id':fields.many2one('stock.location',string='Emplacement',domain=[('usage','=','internal')],required=False,),
        'product_id': fields.many2one('product.product', "Product", domain=[('cost_method','=','average'),('categ_id.cmp_filter','=',True)], required=False,
            help="Product for which recalculation will be done. If not set recalculation will be done for all products with cost method set 'Average Price'"),
    }
    _defaults = {
        'period_id': 60,
    }

    def _get_couple_depot_produit(self,cr,uid,product,location):
        depot_dict=[]
        if not location and not product:
            cr.execute(""" select id from product_loction_cmp order by location_id""")
        elif location and not product:
            cr.execute(""" select id from product_loction_cmp where location_id = %s order by location_id""",(location.id,))
        elif location and  product:
            cr.execute(""" select id from product_loction_cmp where location_id = %s and product_id = %s order by location_id""",(location.id,product.id,))

        couple_depot_produit_dict = cr.dictfetchall()
        couple_depot_produit = [i['id'] for i in couple_depot_produit_dict]
        return couple_depot_produit


    def _get_mouvements(self,cr,uid,operation,sens,period,couple):

        mvt_pool = self.pool.get('stock.move')
        location_pool = self.pool.get('stock.location')

        if operation == 'inventaire':
            id_depot_ref = location_pool.search(cr, uid, [('usage', '=', 'inventory'),('scrap_location','=',False)])
        elif operation == 'achat':
            id_depot_ref = location_pool.search(cr, uid, [('usage', '=', 'supplier')])
        elif operation == 'vente':
            id_depot_ref = location_pool.search(cr, uid, [('usage', '=', 'customer')])
        elif operation == 'transfert':
            id_depot_ref = location_pool.search(cr, uid, [('code', '=', 'TMP')])
        elif operation == 'reclassement':
            id_depot_ref = location_pool.search(cr, uid, [('is_rec', '=', True)])

        # print 'id_depot_ref: ',id_depot_ref
        if id_depot_ref:
            if sens == 'in':
                mvt_ids = mvt_pool.search(cr, uid, [('location_id', '=', id_depot_ref[0]),
                                                   ('location_dest_id', '=', couple.location_id.id),
                                                   ('date_expected', '>=', period.date_start),
                                                   ('date_expected', '<=', period.date_stop),
                                                   ('product_id', '=', couple.product_id.id),
                                                   ('state', '=', 'done')])
            elif sens == 'out':
                mvt_ids = mvt_pool.search(cr, uid, [('location_id', '=', couple.location_id.id),
                                                   ('location_dest_id', '=', id_depot_ref[0]),
                                                   ('date_expected', '>=', period.date_start),
                                                   ('date_expected', '<=', period.date_stop),
                                                   ('product_id', '=', couple.product_id.id),
                                                   ('state', '=', 'done')])
            # print 'Operation : ' + operation + ' ** Sens : '+sens + ' ** Produit : ' + str(couple.product_id.name_template) + ' ** Depot : '+ couple.location_id.name
            # print mvt_ids
            return mvt_ids

    def _get_data_table(self,cr,uid,mvt_ids,sens,operation,couple):
        mvt_ids_obj = self.pool.get('stock.move').browse(cr,uid,mvt_ids)

        coeff = 1
        if sens == 'out':
            coeff = -1
        table_list = []
        for mvt in mvt_ids_obj:

            mvt_id = mvt.id
            op = operation
            ss = sens
            produit = couple.product_id
            depot_id = couple.location_id
            contre_depot_id = depot_id
            contre_produit_id = produit
            if operation == 'transfert':
                if sens == 'out':
                    contre_depot_id = mvt.internal_picking_id.location_dest_id
                else :
                    contre_depot_id = mvt.internal_picking_id.location_src_id
            elif operation == 'reclassement':
                if sens == 'out':
                    contre_produit_id = mvt.reclassification_id.product_dest_id
                else :
                    contre_produit_id = mvt.reclassification_id.product_src_id

            frais = mvt.frais_transfert
            cout_hors_frais = mvt.price_unit

            table = [mvt_id ,
                     operation ,
                     sens,
                     produit.name_template,
                     depot_id.name,
                     contre_depot_id.name,
                     contre_produit_id.name_template
                     ]
            table_list.append(table)
            # print 'table :', table
        return table_list

    def _get_data_frame(self,cr,uid,period,product,location):
        couple_pool = self.pool.get('product.loction.cmp')
        couple_obj = couple_pool.browse(cr, uid, self._get_couple_depot_produit(cr, uid,product,location))

        operation_list = ['inventaire','achat','vente','transfert','reclassement']
        sens_list = ['out','in']

        titre =['mvt_id',
             'operation',
             'sens',
             'produit',
             'depot_id',
             'contre_depot_id',
             'contre_produit_id'
            ]
        df = pd.DataFrame()
        for couple in couple_obj:
            for operation in operation_list:
                for sens in sens_list:
                    mvt_ids = self._get_mouvements(cr,uid,operation,sens,period,couple)
                    if mvt_ids:
                        df = df.append(self._get_data_table(cr,uid,mvt_ids,sens,operation,couple),ignore_index = True)
        df.columns = titre
        return df


    def execute(self,cr,uid,ids,context=None):

        record = self.browse(cr,uid,ids[0])
        period = record.period_id
        product = record.product_id
        location = record.location_id

        df_mvt = self._get_data_frame(cr,uid,period,product,location)
        # couple = df_mvt.query('''operation in ('transfert','reclassement') and sens = 'in' ''')
        filtre = {'operation':['reclassement','transfert','inventaire','achat'],
                  'sens':['in']}
        # couple = df_mvt[df_mvt['operation'].isin(['transfert', 'reclassement']) and df_mvt['sens'] == 'in']
        # colonne = ['operation','sens','produit','depot_id','contre_produit_id','contre_depot_id']
        colonne = ['produit','depot_id','contre_produit_id','contre_depot_id']
        couple = df_mvt[df_mvt[list(filtre)].isin(filtre).all(axis = 1)][list(colonne)].drop_duplicates()
        print couple
        couple.to_csv('/home/holokainf/Desktop/soma_couple.csv',sep=';',header=True,index=False)
        # print couple







cmpu_compute()