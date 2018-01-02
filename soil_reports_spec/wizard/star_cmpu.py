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
from openerp.osv import fields, orm,osv
from datetime import datetime ,timedelta , date
from operator import itemgetter
import numpy as np
import pandas as pd
import calendar as cal
from monthdelta import monthdelta
import time


class cmpu_compute(orm.TransientModel):
    "TODO: Gestion des coulages"
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
        couple_pool = self.pool.get('product.loction.cmp')
        couple_obj = couple_pool.browse(cr,uid,couple_depot_produit)
        return couple_obj

    def _get_df_couple(self,cr,uid,product,location):
    # Attention cette fontion retourne un dataframe
        depot_dict=[]
        if not location and not product:
            cr.execute(""" select id,location_id,product_id from product_loction_cmp order by location_id""")
        elif location and not product:
            cr.execute(""" select id,location_id,product_id from product_loction_cmp where location_id = %s order by location_id""",(location.id,))
        elif location and  product:
            cr.execute(""" select id,location_id,product_id from product_loction_cmp where location_id = %s and product_id = %s order by location_id""",(location.id,product.id,))
        return pd.DataFrame(cr.dictfetchall())

    def _get_mouvements(self,cr,uid,operation,sens,period,couple):

        mvt_pool = self.pool.get('stock.move')
        internal_pool = self.pool.get('product.internal.picking')
        location_pool = self.pool.get('stock.location')

        if operation == 'inventaire':
            id_depot_ref = location_pool.search(cr, uid, [('usage', '=', 'inventory'),('scrap_location','=',False)])
        elif operation == 'achat':
            id_depot_ref = location_pool.search(cr, uid, [('usage', '=', 'supplier')])
        elif operation == 'vente':
            id_depot_ref = location_pool.search(cr, uid, [('usage', '=', 'customer')])
        elif operation == 'transfert':
            id_depot_ref = location_pool.search(cr, uid, [('code', '=', 'TMP')])
            # if sens == 'coulage':
            #     depot_src_id = location_pool.search(cr, uid, [('code', '=', 'TMP')])
            #     depot_dest_id = location_pool.search(cr, uid, [('for_lost', '=', True)])
            #     internal_picking_id = internal_pool.search(cr,uid,[('location_src_id','=',couple.location_id.id),
            #                                                        ('product_id','=',couple.product_id.id)])
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
            # elif sens == 'coulage':
            #     print 'coucoucocuocu'
            #     mvt_ids = mvt_pool.search(cr, uid, [('location_id', '=', depot_src_id[0]),
            #                                        ('location_dest_id', '=', depot_dest_id[0]),
            #                                        ('date_expected', '>=', period.date_start),
            #                                        ('date_expected', '<=', period.date_stop),
            #                                        ('product_id', '=', couple.product_id.id),
            #                                        ('state', '=', 'done'),('internal_picking_id','=',internal_picking_id)])

            # print 'Operation : ' + operation + ' ** Sens : '+sens + ' ** Produit : ' + str(couple.product_id.name_template) + ' ** Depot : '+ couple.location_id.name
            # print mvt_ids
            return mvt_ids

    def _get_data_table(self,cr,uid,mvt_ids,sens,operation,couple):

        def _get_depot(operation, sens, mvt, depot_id):
            if operation == 'transfert':
                if sens in ['out','coulage']:
                    contre_depot_id = mvt.internal_picking_id.location_dest_id.id
                else:
                    contre_depot_id = mvt.internal_picking_id.location_src_id.id
            elif operation == 'reclassement':
                contre_depot_id = depot_id.id
            else:
                contre_depot_id = np.nan
            return contre_depot_id

        def _get_product(operation, sens, mvt, produit):
            if operation == 'transfert':
                contre_produit_id = produit.id
            elif operation == 'reclassement':
                if sens == 'out':
                    contre_produit_id = mvt.reclassification_id.product_dest_id.id
                else:
                    contre_produit_id = mvt.reclassification_id.product_src_id.id
            else:
                contre_produit_id = np.nan
            return contre_produit_id

        def _get_quantity(operation, sens, mvt):
            if operation == 'achat':
                return mvt.converted_quantity
            else:
                return mvt.product_qty

        def _get_frais(operation, sens, mvt):
            if operation == 'achat':
                return mvt.frait_achat * mvt.product_qty / mvt.converted_quantity
            else:
                return mvt.frais_transfert

        def _cost_without_taxe(operation, sens, mvt):
            if operation == 'achat':
                return mvt.price_unit_net * mvt.product_qty / mvt.converted_quantity
            else:
                return mvt.price_unit_net

        def _get_reclassification_id(operation, mvt):
            if operation ==  "reclassement":
                return mvt.reclassification_id.id
            else:
                return  np.nan

        def _get_transfert_id(operation, mvt):
            if operation == "transfert":
                return mvt.internal_picking_id.id
            else:
                return np.nan

        mvt_ids_obj = self.pool.get('stock.move').browse(cr,uid,mvt_ids)

        coeff = 1
        if sens == 'out':
            coeff = -1

        table_list = []
        for mvt in mvt_ids_obj:
             table = [mvt.id,
                     operation ,
                     sens,
                     couple.product_id.id,
                     couple.location_id.id,
                     _get_depot(operation, sens, mvt, couple.location_id),
                     _get_product(operation,sens,mvt,couple.product_id),
                     _get_quantity(operation, sens, mvt) * coeff,
                     _get_frais(operation, sens, mvt),
                     _cost_without_taxe(operation, sens, mvt),
                      _get_reclassification_id(operation, mvt),
                      _get_transfert_id(operation,mvt)
                      ]
             table_list.append(table)
        return table_list

    def _get_data_frame(self,cr,uid,period,product,location):
        couple_pool = self.pool.get('product.loction.cmp')
        couple_obj = self._get_couple_depot_produit(cr, uid,product,location)

        operation_list = ['inventaire','achat','vente','transfert','reclassement']
        sens_list = ['out','in']

        titre =['mvt_id',
            'operation',
            'sens',
            'produit',
            'depot_id',
            'contre_depot_id',
            'contre_produit_id',
            'quantity',
            'frais',
            'cout_hors_frais','reclassement_id','transfert_id']
        df = pd.DataFrame()
        for couple in couple_obj:
            for operation in operation_list:
                for sens in sens_list:
                    mvt_ids = self._get_mouvements(cr,uid,operation,sens,period,couple)
                    if mvt_ids:
                        df = df.append(self._get_data_table(cr,uid,mvt_ids,sens,operation,couple),ignore_index = True)
        df.columns = titre
        df['total_frais'] = df['frais']*df['quantity']
        df['total_cout_hors_frais'] = df['cout_hors_frais'] * df['quantity']
        df['total_cout'] = df['total_cout_hors_frais'] + df['total_frais']

        return df

    def _get_depth(self,cr,uid,df_mvt):

        def liste_init(x):
            val = []
            if not pd.isnull(x['contre_produit_id']):
                tup = (x['contre_produit_id'], x['contre_depot_id'])
                val.append(tup)
            return val

        def fct(x):
            val = [] + x['trajet']
            if not pd.isnull(x['produit_src']):
                tup = (x['produit_src'], x['depot_src'])
                val.append(tup)
            return val

        def deep(c):
            while c[c['contre_produit_id'] != np.nan].count()['contre_produit_id'] != 0:
                c = pd.merge(c, cp, how='left', left_on=['contre_produit_id', 'contre_depot_id'],
                             right_on=['produit_dest', 'depot_dest'])
                c['trajet'] = c.apply(lambda x: fct(x), axis=1)
                c['contre_produit_id'] = c['produit_src']
                c['contre_depot_id'] = c['depot_src']
                c.drop(['produit_dest', 'depot_dest', 'produit_src', 'depot_src'], axis=1, inplace=True)
            return c

        # couple = df_mvt.query('''operation in ('transfert','reclassement') and sens = 'in' ''')
        filtre = {'operation':['reclassement','transfert','inventaire','achat'],
                  'sens':['in']}

        colonne = ['produit','depot_id','contre_produit_id','contre_depot_id']
        couple = df_mvt[df_mvt[list(filtre)].isin(filtre).all(axis = 1)][list(colonne)].drop_duplicates()
        couple.to_csv('/home/holokainf/Desktop/soma_couple.csv',sep=';',header=True,index=False)
        cp = couple.copy()
        cp.columns = ['produit_dest', 'depot_dest', 'produit_src', 'depot_src']
        couple['trajet'] = couple.apply(lambda x: liste_init(x), axis=1)

        dp = deep(couple)
        dp['profondeur'] = dp.apply(lambda x: len(x['trajet']) + 1, axis=1)
        dp.drop(['contre_produit_id', 'contre_depot_id'], axis=1, inplace=True)
        dp.sort_values(['produit', 'depot_id'])
        dp.to_csv('/home/holokainf/Desktop/soma_deep.csv', sep=';', header=True, index=False)
        return dp.groupby(['produit', 'depot_id'], sort=False)['profondeur'].max().reset_index().sort_values(['profondeur'])

    def _get_cmpu_hist(self,cr,uid,period_id,product_id,location_id):
        #TODO: si pas de cpmu N-1 alors le créer et fixer les valeurs à 0
        hist_pool = self.pool.get('month.stock.qty')
        hist_ids = hist_pool.search(cr,uid,[('period_id','=',period_id),('product_id','=',product_id),('location_id','=',location_id)])
        if hist_ids:
            return  hist_pool.browse(cr,uid,hist_ids[0])
        else:
            hist_ids = hist_pool.create(cr,uid,{'period_id':period_id,'location_id':location_id,'product_id':product_id,'cmp':0,'end_qty':0})
            return hist_pool.browse(cr, uid, hist_ids)

    def _get_previous_period(self,cr,uid,period):
        period_pool = self.pool.get('account.period')
        period_datetime = datetime.strptime(period.date_stop,'%Y-%m-%d')- monthdelta(1)
        print 'period_datetime: ',period_datetime
        previous_period_range = cal.monthrange(period_datetime.year,period_datetime.month)
        print previous_period_range
        date_stop = datetime(period_datetime.year,period_datetime.month, previous_period_range[1]).strftime('%Y-%m-%d')
        print 'date_stop: ',date_stop
        previous_period_id = period_pool.search(cr,uid,[('date_stop','=',date_stop)])
        print 'previous_period_id: ',previous_period_id
        previous_period = period_pool.browse(cr,uid,previous_period_id[0])
        # print previous_period
        return previous_period

    def _cost_update_mvt_out(self,cr,uid,df_mvt,cmpu,depot_id,produit_id):
         filtre1 = ['inventaire', 'transfert', 'reclassement']
         filtre2 = ['vente']
         for row in df_mvt.itertuples():
             if row.depot_id == depot_id and row.produit == produit_id:
                 if row.sens == 'out' and row.operation in filtre1:
                     df_mvt.loc[row.Index, 'cout_hors_frais'] = cmpu
                     df_mvt.loc[row.Index, 'total_cout_hors_frais'] = cmpu * row.quantity
                     df_mvt.loc[row.Index, 'total_cout'] = (cmpu + row.frais) * row.quantity
                 elif row.operation in filtre2:
                     df_mvt.loc[row.Index, 'cout_hors_frais'] = cmpu
                     df_mvt.loc[row.Index, 'total_cout_hors_frais'] = cmpu * row.quantity
                     df_mvt.loc[row.Index, 'total_cout'] = (cmpu + row.frais) * row.quantity

    def _cost_update_mvt_in(self,cr,uid,df_mvt,depot_id,produit_id):
        filtre1 = ['transfert', 'reclassement']
        for row in df_mvt.itertuples():
            if row.depot_id == depot_id and row.produit == produit_id and row.sens == 'in' and row.operation in filtre1:
                # print 'sisisisis: ',row.operation
                if row.operation  == 'transfert':
                    print 'transfert_id: ',row.mvt_id
                    if df_mvt['cout_hors_frais'][(df_mvt.transfert_id == row.transfert_id) & (df_mvt.sens == 'out')].size !=0:
                        cmpu = df_mvt['cout_hors_frais'][(df_mvt.transfert_id == row.transfert_id) & (df_mvt.sens == 'out')].item() + df_mvt['frais'][
                            (df_mvt.transfert_id == row.transfert_id) & (df_mvt.sens == 'out')] .item()
                    else:
                        print 'Id origine non dans dft mouvement'
                        stock_pool = self.pool.get('stock.move')
                        stock_ids = stock_pool.search(cr,uid,[('internal_picking_id','=',row.transfert_id),('location_dest_id','=',int(row.depot_id))])
                        #TODO: rajouter exception si aucun id non trouvé et si plusieurs ids trouveé
                        stock_obj = stock_pool.browse(cr,uid,stock_ids[0])
                        cmpu = stock_obj.price_unit + stock_obj.frais_transfert
                    print 'transfert:',cmpu
                elif row.operation == 'reclassement':
                    print 'Reclassement_id: ',row.mvt_id
                    cmpu_out = df_mvt['cout_hors_frais'][(df_mvt.reclassement_id == row.reclassement_id) & (df_mvt.sens == 'out')].item() + df_mvt['frais'][
                        (df_mvt.reclassement_id == row.reclassement_id) & (df_mvt.sens == 'out')] .item()
                    qty_out = df_mvt['quantity'][(df_mvt.reclassement_id == row.reclassement_id) & (df_mvt.sens == 'out')].item()
                    cmpu = abs(cmpu_out * qty_out / row.quantity)
                    print 'Reclassement:',cmpu

                df_mvt.loc[row.Index, 'cout_hors_frais'] = cmpu
                df_mvt.loc[row.Index, 'total_cout_hors_frais'] = cmpu * row.quantity
                df_mvt.loc[row.Index, 'total_cout'] = (cmpu + row.frais) * row.quantity

    def _get_sum_entry(self,cr,uid,df_mvt,row):
        # 1.2 on récupère le total des entrées modifiants le couts
        df_cout = df_mvt[(
                             (df_mvt.operation == 'achat')
                             | ((df_mvt.operation == 'inventaire') & (df_mvt.sens == 'in'))
                             | ((df_mvt.operation == 'transfert') & (df_mvt.sens == 'in'))
                             | ((df_mvt.operation == 'reclassement') & (df_mvt.sens == 'in'))
                         )
                         & (df_mvt.depot_id == row.depot_id) & (df_mvt.produit == row.produit)]


        return {'total_qty': df_cout['quantity'].sum(),
                'total_cost' : df_cout['total_cout'].sum()}

    def _set_cmpu_history(self,cr,uid,period_id,depot_id,product_id,cmpu,end_qty):
        cmpu_hist_pool = self.pool.get('month.stock.qty')
        cmpu_hist_id = self._get_cmpu_hist(cr,uid,period_id,product_id,depot_id)
        if cmpu_hist_id:
            # cmpu_hist_obj = cmpu_his_pool.browse(cr,uid,cmpu_hist_id[0])
            cmpu_hist_pool.write(cr,uid,cmpu_hist_id.id,{'cmp':cmpu,'end_qty':end_qty})
        else:
            cmpu_hist_pool.create(cr,uid,{'period_id':period_id,'location_id':depot_id,'product_id':product_id,'cmp':cmpu,'end_qty':end_qty})

    def _unlink_cmpu_hist(self,cr,uid,period):
        hist_obj = self.pool.get('month.stock.qty')
        hist_ids = hist_obj.search(cr,uid,[('period_id','=',period.id)])
        hist_obj.unlink(cr,uid,hist_ids)

    def _update_stock_move(self,cr,uid,df_mvt):
        sm_pool = self.pool.get('stock.move')
        d = df_mvt[df_mvt['operation'].isin(['vente', 'transfert', 'reclassement'])]
        d = d[['mvt_id','frais','cout_hors_frais']]
        d.columns = ['id','frais_transfert','price_unit_net']
        d['price_unit'] = d['frais_transfert'] + d['price_unit_net']
        # sm_ids  = d['mvt_id'].tolist()
        # val = d.to_dict('records')
        for row in d.itertuples():
            sm_pool.write(cr,uid, row.id, {'price_unit_net': row.price_unit_net, 'price_unit': row.price_unit})

    def _update_account_move_line(self,cr,uid,df_mvt):
        #TODO: Prise en compte des pièce comptable de coulage
        am_pool  = self.pool.get('account.move')
        aml_pool = self.pool.get('account.move.line')
        sm_pool = self.pool.get('stock.move')
        pp_pool = self.pool.get('product.product')

        #1.1 On recupère la liste des mvt_ids
        d = df_mvt[df_mvt['operation'].isin(['vente', 'transfert', 'reclassement'])]
        sm_ids = d['mvt_id'].tolist()
        print 'Nombre de mouvement: ', len(sm_ids)

        #1.2. On récupère la lise des aml_ids
        aml_ids = aml_pool.search(cr,uid,[('stock_move_id','in',sm_ids)])
        print 'aml_ids: ',aml_ids
        aml_obj = aml_pool.browse(cr,uid,aml_ids)
        for aml in aml_obj:
            # ('state', '=', 'draft')
            if aml.move_id.state == 'posted':
                # am_pool.write(cr,uid,aml.move_id.id,{'state':'draft'})
                am_pool.button_cancel(cr, uid,[aml.move_id.id])
            #2.1 Operation détermine compte de contrepartie
            operation = d['operation'][d.mvt_id == aml.stock_move_id.id].item()

            #2.1 Sens détermine compte de contrepartie
            sens = d['sens'][d.mvt_id == aml.stock_move_id.id].item()

            # 2.3 On récupére les comptes
            account_value_id = aml.stock_move_id.product_id.categ_id.property_stock_valuation_account_id.id
            if operation == 'vente':
                account_counterpart_id = aml.stock_move_id.product_id.property_stock_account_output.id
                if not account_counterpart_id:
                    account_counterpart_id = \
                        aml.stock_move_id.product_id.categ_id.property_stock_account_output_categ.id
            elif operation == 'transfert':
                account_counterpart_id = aml.stock_move_id.product_id.stock_tmp_account_input.id
                if not account_counterpart_id:
                    account_counterpart_id = aml.stock_move_id.product_id.categ_id.stock_tmp_account_input_categ.id
            elif operation == 'reclassement':
                account_counterpart_id = None

            # 2.4. On vérifie le compte
            if sens == 'out':
                if aml.account_id.id == account_value_id:
                    print 'lououlouuo'
                    montant = abs(d['total_cout_hors_frais'][d.mvt_id == aml.stock_move_id.id].item())
                    aml_pool.write(cr,uid,aml.id,{'debit': 0,'credit' : montant})
                elif aml.account_id.id == account_counterpart_id:
                    montant = abs(d['total_cout'][d.mvt_id == aml.stock_move_id.id].item())
                    aml_pool.write(cr,uid,aml.id,{'debit': montant,'credit' : 0})
            elif sens == 'in':
                if aml.account_id.id == account_value_id:
                    montant = abs(d['total_cout'][d.mvt_id == aml.stock_move_id.id].item())
                    aml_pool.write(cr,uid,aml.id,{'debit': montant,'credit' : 0})
                elif aml.account_id.id == account_counterpart_id:
                    montant = abs(d['total_cout_hors_frais'][d.mvt_id == aml.stock_move_id.id].item())
                    aml_pool.write(cr,uid,aml.id,{'debit': 0,'credit' : montant})

        #1.3.Quel traitememnt a effectuer pour chaque type d'opération

        # accounts = self.pool.get('product.product').get_product_accounts(cr, uid, _product.id, context)

            print 'PC_id: ',aml.move_id.id, ' *** AML_id: ',aml.id,' *** Move_id: ',aml.stock_move_id.id
            print ' *** Cpt: ',aml.account_id.id,' *** Cptt_art: ', account_value_id , ' *** ',account_counterpart_id, ' Montant:', montant

    def execute(self,cr,uid,ids,context=None):

        debut = time.time()
        record = self.browse(cr,uid,ids[0])
        period = record.period_id
        product = record.product_id
        location = record.location_id

        #1.1. Récupère la liste des mouvements du mois retraité
        df_mvt = self._get_data_frame(cr,uid,period,product,location)
        df_mvt.to_csv('/home/holokainf/Desktop/soma_liste_mvt.csv', sep=';', header=True, index=False)

        #1.2. Récupère la liste des dépôts par ordre de calcul
        col = ['mvt_id','operation','sens','produit','depot_id','contre_depot_id','contre_produit_id']
        profondeur = self._get_depth(cr,uid,df_mvt[list(col)][df_mvt.sens != 'coulage'])
        profondeur.to_csv('/home/holokainf/Desktop/soma_profondeur.csv', sep=';', header=True, index=False)

        #1.3.On récupère la period_id du mois précedent
        previous_period = self._get_previous_period(cr, uid, period)

        #1.4. on suprime toute l'historique de la période
        self._unlink_cmpu_hist(cr, uid, period)

        #1.5. On recréer toute l'historique à partir du mois N-1
        couple_df = self._get_df_couple(cr, uid, product, location)
        for _, row in couple_df.iterrows():
            previous_cmpu  = self._get_cmpu_hist(cr,uid,previous_period.id,row['product_id'],row['location_id'])
            self._set_cmpu_history(cr, uid, period.id, row['location_id'], row['product_id'], previous_cmpu.cmp,previous_cmpu.end_qty)


        for _,row in profondeur.iterrows():

            # 5.1.On récupère la liste des CMPU du mois passé comme stock et cmpu initiale
            #TODO: Voir problème si CMPU N-1 non trouvé?
            previous_cmpu = self._get_cmpu_hist(cr,uid,previous_period.id,row['produit'],row['depot_id'])
            # print 'previous_cmpu: ',previous_cmpu
            print 'depot_id: ',row['depot_id'],' *** Produit_id: ',row['produit']

            #5.2. On récupère la total des coûts des achats,inventaires et transferts entrants.
            sum_mvt_period = self._get_sum_entry(cr,uid,df_mvt,row)

            #5.3 On calcul le CMPU du couple
            cmpu_stock = previous_cmpu.end_qty + sum_mvt_period.get('total_qty')
            cmpu_cost = previous_cmpu.cmp * previous_cmpu.end_qty + sum_mvt_period.get('total_cost')
            if cmpu_stock == 0:
                cmpu = 0
            else:
                cmpu = cmpu_cost / cmpu_stock
            end_qty = df_mvt[(df_mvt.depot_id == row.depot_id) & (df_mvt.produit == row.produit)]['quantity'].sum() + previous_cmpu.end_qty

            if row['produit'] == 4968 and row['depot_id'] ==259:
                print 'cmpu_stock: ', cmpu_stock, ' cmpu_cost:',cmpu_cost, 'cmpu:',cmpu, 'end_qty:',end_qty

            #5.4. On actualise la tables des sorties:
            self._cost_update_mvt_out(cr, uid, df_mvt, cmpu,row['depot_id'],row['produit'])

            #5.5. On actualise les entrée provenant de ce couple
            self._cost_update_mvt_in(cr, uid, df_mvt,row['depot_id'], row['produit'])

            #5.6 on actualise ou on crée  la ligne de CMPU pour le mois calculé
            self._set_cmpu_history(cr, uid, period.id, row['depot_id'], row['produit'], cmpu, end_qty)

        #1.6 On actualise la table stock la table stock move
        self._update_stock_move(cr,uid,df_mvt)
        df_mvt.to_csv('/home/holokainf/Desktop/soma_liste_mvt_final.csv', sep=';', header=True, index=False)

        # self._update_account_move_line(cr, uid, df_mvt)

        fin = time.time()
        print (fin - debut)/60



cmpu_compute()