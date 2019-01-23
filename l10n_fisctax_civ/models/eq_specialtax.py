# -*- coding: utf-8 -*-
from datetime import datetime, timedelta
import time
from openerp.osv import osv, fields, orm
from openerp.tools.translate import _
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT, DEFAULT_SERVER_DATETIME_FORMAT
import openerp.addons.decimal_precision as dp
from openerp import workflow



# declaration de la classe tse  dans fiscalité
class eq_specialtax_declaration(osv.osv): 
    _name = 'eq.specialtax.declaration'
    _res_name='name'

    #Début de la déclaration des fonctions de calcul de TSE ou EQ SPECIAL TAX

    def on_change_company_elements(self, cr, uid, ids, company_id, context=None):
        values = {}
        if company_id:
            company = self.pool.get('res.company').browse(cr, uid, company_id, context=context)
            values = {
                'tax_period':company.impot_period,
                'tax_service':company.service,
                'company_tax_code':company.vat,
            }
        return {'value': values}

    def taxable_revenu_function(self, cr, uid, ids):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] =record.revenu_wout_tax

        return x

    def tax_amount_calcul(self, cr, uid, ids):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] =(record.taxable_revenu * record.rate)/100

        return x

    #fin de la déclaration des fonctions de calcul de TSE ou EQ SPECIAL TAX
    _columns = {
        'name': fields.char('Name'),
        'state':fields.selection([('draft','Draft'),('validation','Validation'),('termined','Termined')],'State',required=True,readonly=True),
        'date': fields.date('Date', required=True),
        'periodicity':fields.selection([('monthly','Monthly'),('termly','Termly')],'Periodicity', required=True),
        'tax_period':fields.char("Tax Period"),
        'tax_service':fields.char("Tax Service"),
        'company_id':fields.many2one('res.company',"Company Name",ondelete='set null', track_visibility='onchange',
            select=True, help="Linked company (optional). Usually created when converting the eq_specialtax.", required=True),
        'company_tax_code':fields.char('Tax ID'),
        'revenu_wout_tax':fields.float("Total Revenu Without Tax"),
        'ops_exempt_tax_revenu':fields.float("Exempt Tax Operations (Oil Products) to deduct"),
        'delivery_onself_revenu': fields.float('Delivery Oneself Revenu to deduct'),
        'taxable_revenu':fields.function(taxable_revenu_function,string="Taxable Revenu", method=True,type='float', onchange=True),
        'rate':fields.float('Rate (In %)'),
        'tax_amount':fields.function(tax_amount_calcul,string='Tax Amount', method=True,type='float', onchange=True),
        'regulation':fields.float('Regularisation'),
        'amount_topay':fields.function(tax_amount_calcul,string='Tax Amount to pay', method=True,type='float', onchange=True),
        'amount_toreport':fields.float('Tax Amount to report'),
        
    }
    _defaults = {
    'active': True, 
    'state': 'draft',
    'name':lambda self,cr,uid,context={}: self.pool.get('ir.sequence').get(cr, uid, 'eq.specialtax.declaration'),
    }
