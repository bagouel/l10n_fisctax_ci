# -*- coding: utf-8 -*-
from datetime import datetime, timedelta
import time
from openerp.osv import osv, fields, orm
from openerp.tools.translate import _
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT, DEFAULT_SERVER_DATETIME_FORMAT
import openerp.addons.decimal_precision as dp
from openerp import workflow



# declaration de la classe its  dans fiscalité
class wagetax_declaration(osv.osv): 
    _name = 'wagetax.declaration'
    _res_name='name'

    # Debut des fonctions de calcul de ITS OU WAGETAX
    def on_change_company_elements(self, cr, uid, ids, company_id, context=None):
        values = {}
        if company_id:
            company = self.pool.get('res.company').browse(cr, uid, company_id, context=context)
            values = {
                'tax_period':company.impot_period,
                'tax_service':company.service,
                'company_tax_code':company.vat,
                'initials':company.sigle,
                'company_goal':company.objet,
                'street':company.street2,
                'pobox':company.zip,
                'phone':company.phone,
                'street2':company.street,
                'district':company.quartier,
                'email':company.email,
            }
        return {'value': values}


    def taxamount_1_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.amount_tv1 - record.amount_exo1

        return x

    def taxamount_2_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.amount_tv2 - record.amount_exo2

        return x

    def taxamount_3_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.amount_tv3 - record.amount_exo3

        return x

    def taxamount_4_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.amount_tv4 - record.amount_exo4

        return x

    def revenu_ni1_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.amount_tax1 - ((record.amount_tax1 * record.taxreduction_1)/100)

        return x

    def revenu_ni2_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.amount_tax2 - ((record.amount_tax2 * record.taxreduction_2)/100)

        return x

    def revenu_ni3_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.amount_tax3 - ((record.amount_tax3 * record.taxreduction_3)/100)

        return x

    def revenu_ni4_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.amount_tax4 - ((record.amount_tax4 * record.taxreduction_4)/100)

        return x

    def amount_tb_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.amount_rv + record.amount_an + record.amount_a

        return x

    def wageamount_2_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.amount_tb - record.wageamount_1 - record.totalrevenu_1

        return x
        
    def allowance_2_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.amount_tb - record.allowance_1 - record.netrevenu_1

        return x
        
    def totalrevenu_2_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = (record.wageamount_2 * record.annual_tax ) / 100

        return x

    def netrevenu_2_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = (record.allowance_2 * record.annual_tax ) / 100

        return x

    def totalnetamount_1_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.totalrevenu_2 +( record.revenu_ni1 + record.revenu_ni2) + (record.revenu_ni3 + record.revenu_ni4)

        return x


    def totalnetamount_2_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.netrevenu_2 +( record.revenu_ni1 + record.revenu_ni2) + (record.revenu_ni3 + record.revenu_ni4)

        return x

    def totalnetamount_1_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.totalrevenu_2 +( record.revenu_ni1 + record.revenu_ni2) + (record.revenu_ni3 + record.revenu_ni4)

        return x

    def taxamount_1(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = (record.base_1 * record.taxrate_1 ) / 100

        return x

    def taxamount_2(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = (record.base_2 * record.taxrate_2) / 100

        return x

    def taxamount_3(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = (record.base_3 * record.taxrate_3 ) / 100

        return x

    def total_wr_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.taxamount_1 + record.taxamount_2 + record.taxamount_3

        return x

    def amount_e1_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = (record.revenu_nii1 * record.rate_e1 ) / 100

        return x

    def amount_e2_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = (record.revenu_nii2 * record.rate_e2 ) / 100

        return x

    def amount_e3_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = (record.revenu_nii3 * record.rate_e3 ) / 100

        return x

    def amount_e4_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = (record.revenu_nii4 * record.rate_e4 ) / 100

        return x

    def amount_e5_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = (record.revenu_nii5 * record.rate_e5 ) / 100

        return x

    def total_contribution_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.amount_e1 + record.amount_e2 + record.amount_e3 + record.amount_e4 + record.amount_e5

        return x

    def amount_retained_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = (record.revenu_netimp * record.rate_re ) / 100

        return x

    def total_r_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.tax + record.contribution_n + record.tax_gsr + record.contribution_e + record.contribution_nce

        return x

    def amount_tp_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.total_r + record.amount_retained + record.totalcontribution 
            if x[record.id]>=0 :
                x[record.id] = x[record.id]
            else:
                x[record.id]= 0

        return x

    def amount_tr_func(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.total_r + record.amount_retained + record.totalcontribution 
            if x[record.id]<0 :
                x[record.id] = 0 - x[record.id]
            else:
                x[record.id]= 0

        return x
    # Fin des fonctions de calcul de ITS OU WAGETAX


    _columns = {
        'name': fields.char('Name'),
        'state':fields.selection([('draft','Draft'),('validation','Validation'),('termined','Termined')],'State',required=True,readonly=True),
        'date': fields.date('Date', required=True),
        'periodicity':fields.selection([('monthly','Monthly'),('termly','Termly')],'Periodicité', required=True),
        'tax_period':fields.char("Tax Period"),
        'tax_service':fields.char("Tax Service"),
        'company_tax_code':fields.char('Tax ID'),
        'initials':fields.char('Initials'),
        'company_goal':fields.char('Company Goal'),
        'street': fields.char('Street'),
        'pobox':fields.char("P.O Box"),
        'phone':fields.char("Phone"),
        'street2':fields.char("Street 2"),
        'district':fields.char("District"),
        'company_id':fields.many2one('res.company',"Company Name",ondelete='set null', track_visibility='onchange',
            select=True, help="Linked company (optional). Usually created when converting the wagetax.", required=True),
        'email':fields.char('Email'),

        # Determination de l'assiette
        'annual_tax':fields.integer('Annual Tax ( in %)'),
            # 1er tableau
        'amount_tv1':fields.float(''),
        'workforce_1':fields.integer(""),
        'amount_exo1':fields.float(""),
        'amount_tax1':fields.function(taxamount_1_func, method=True,type='float', onchange=True),
        'taxreduction_1':fields.integer(""),
        'revenu_ni1':fields.function(revenu_ni1_func, method=True,type='float', onchange=True),
        'amount_tv2':fields.float(''),
        'workforce_2':fields.integer(""),
        'amount_exo2':fields.float(""),
        'amount_tax2':fields.function(taxamount_2_func, method=True,type='float', onchange=True),
        'taxreduction_2':fields.integer(""),
        'revenu_ni2':fields.function(revenu_ni2_func, method=True,type='float', onchange=True),

            # 2e tableau
        'amount_tv3':fields.float(''),
        'workforce_3':fields.integer(""),
        'amount_exo3':fields.float(""),
        'amount_tax3':fields.function(taxamount_3_func, method=True,type='float', onchange=True),
        'taxreduction_3':fields.integer(""),
        'revenu_ni3':fields.function(revenu_ni3_func, method=True,type='float', onchange=True),
        'amount_tv4':fields.float(''),
        'workforce_4':fields.integer(""),
        'amount_exo4':fields.float(""),
        'amount_tax4':fields.function(taxamount_4_func, method=True,type='float', onchange=True),
        'taxreduction_4':fields.integer(""),
        'revenu_ni4':fields.function(revenu_ni4_func, method=True,type='float', onchange=True),

           # 3e tableau
        'amount_rv':fields.float(''),
        'amount_an':fields.float(""),
        'amount_a':fields.float(""),
        'amount_tb':fields.function(amount_tb_func, string="Gross total revenu" ,method=True,type='float', onchange=True),

           # 4e tableau
        'wageamount_1':fields.float(''),
        'allowance_1':fields.float(""),
        'totalrevenu_1':fields.float(""),
        'netrevenu_1':fields.float(""),
        'wageamount_2':fields.function(wageamount_2_func, method=True,type='float', onchange=True),
        'allowance_2':fields.function(allowance_2_func, method=True,type='float', onchange=True),
        'totalrevenu_2':fields.function(totalrevenu_2_func, method=True,type='float', onchange=True),
        'netrevenu_2':fields.function(netrevenu_2_func, method=True,type='float', onchange=True),

           # 5e tableau
        'totalnetamount_1':fields.function(totalnetamount_1_func, method=True, type='float', onchange=True),
        'totalnetamount_2':fields.function(totalnetamount_2_func, method=True, type='float', onchange=True),

        # Détermination de l'impot

        # impots retenus sur  les salaires
            # 1er tableau
        'base_1':fields.float(""),
        'rateimpot1':fields.float(""),
        'taxrate_1':fields.integer(""),
        'taxamount_1':fields.function(taxamount_1, method=True, type='float', onchange=True),
        'base_2':fields.float(''),
        'taxrate_2':fields.integer(""),
        'taxamount_2':fields.function(taxamount_2, method=True, type='float', onchange=True),
        'base_3':fields.float(""),
        'taxrate_3':fields.integer(""),
        'taxamount_3':fields.function(taxamount_3, method=True, type='float', onchange=True),

            # 2e tableau
        'total_wr':fields.function(total_wr_func, string='Total wage retenues', method=True, type='float', onchange=True),
        # contributions a la charge de  l'employeur
            # 1er tableau
        'wageforce_e1':fields.integer(""),
        'revenu_nii1':fields.float(""),
        'rate_e1':fields.integer(""),
        'amount_e1':fields.function(amount_e1_func, method=True, type='float', onchange=True),
        'wageforce_e2':fields.integer(""),
        'revenu_nii2':fields.float(""),
        'rate_e2':fields.integer(""),
        'amount_e2':fields.function(amount_e2_func, method=True, type='float', onchange=True),
        'workforce_e3':fields.integer(""),
        'revenu_nii3':fields.float(""),
        'rate_e3':fields.integer(""),
        'amount_e3':fields.function(amount_e3_func, method=True, type='float', onchange=True),
        'workforce_e4':fields.integer(""),
        'revenu_nii4':fields.float(""),
        'rate_e4':fields.integer(""),
        'amount_e4':fields.function(amount_e4_func, method=True, type='float', onchange=True),
        'workforce_e5':fields.integer(""),
        'revenu_nii5':fields.float(""),
        'rate_e5':fields.integer(""),
        'amount_e5':fields.function(amount_e5_func, method=True, type='float', onchange=True),
            # 2e tableau
        'totalcontribution':fields.function(total_contribution_func, string="Employer total contribution", method=True, type='float', onchange=True),

            # 1er tableau ( impots retenu sur salaire)
            # retenue specifique du regime forestier en cas de fermage
        'revenu_netimp':fields.float(""),
        'rate_re':fields.integer(""),
        'amount_retained':fields.function(amount_retained_func ,method=True, type='float', onchange=True),

            # regularisation

             #1er tableau
        'tax':fields.float(""),
        'contribution_n':fields.float(""),
        'tax_gsr':fields.float(""),
        'contribution_e':fields.float(""),
        'contribution_nce':fields.float(""),
        'total_r':fields.function(total_r_func ,method=True, type='float', onchange=True),
             # 2e tableau
        'amount_tp':fields.function(amount_tp_func, string="Amount to pay", method=True, type='float', onchange=True),
        'amount_tr':fields.function(amount_tr_func, string="Amount to report", method=True, type='float', onchange=True),
    }
    _defaults = {
    'active': True, 
    'state': 'draft',
    'name':lambda self,cr,uid,context={}: self.pool.get('ir.sequence').get(cr, uid, 'wagetax.declaration'),
    }
