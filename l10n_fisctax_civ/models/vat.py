# -*- coding: utf-8 -*-
from datetime import datetime, timedelta
import time
from openerp.osv import osv, fields, orm
from openerp.tools.translate import _
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT, DEFAULT_SERVER_DATETIME_FORMAT
import openerp.addons.decimal_precision as dp
from openerp import workflow



# declaration de la classe tva dans fiscalité
class vat_decalaration(osv.osv): 
    _name = 'vat.declaration'
    _res_name='name'

    #Début de la déclaration des fonctions de calcul de TVA ou VAT

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

    def on_change_deductible_vat_statement(self, cr, uid, ids, deductible_id, context=None):
        values = {}
        if name :
            deductible = self.pool.get('deductiblevat.statement').browse(cr, uid, name, context=context)
            values = {
                'deductible_vat_total':deductible.total_statement,
            }
        return {'value': values}

    def taxable_revenu_function(self, cr, uid, ids, context=None):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] =record.revenu_wout_tax

        return x

    def normal_rate_revenu_amount_calcul(self, cr, uid, ids, context=None):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] =(record.normal_rate_vat_amount * record.normal_rate)/100

        return x

    def minimal_rate_revenu_amount_calcul(self, cr, uid, ids, context=None):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] =(record.minimal_rate_vat_amount * record.minimal_rate)/100

        return x
    # Fin de la declaration des fonctions de calcul de TVA ou VAT

    _columns = {
        'name': fields.char('Name'),
        'state':fields.selection([('draft','Draft'),('validation','Validation'),('termined','Termined')],'State',required=True,readonly=True),
        'date': fields.date('Date',required=True),
        'periodicity':fields.selection([('monthly','Monthly'),('termly','Termly')],'Periodicity',required=True),
        'tax_period':fields.char('Tax Period'),
        'company_id':fields.many2one('res.company',"Company Name",ondelete='set null', track_visibility='onchange',
            select=True, help="Linked company (optional). Usually created when converting the VAT.", required=True),
        'tax_service':fields.char("Tax Service"),
        'company_tax_code':fields.char('Tax ID'),
        # opeérations réalisees
        'revenu_wout_vat':fields.float("Total revenu without tax"),
        'export_deduction':fields.float("Export deduction"),
        'legal_ops_exempt_vat_revenu':fields.float("Legal Exempt operations to deduct"),
        'conv_ops_exempt_vat_revenu':fields.float("Conventional Exempt operation to deduct (join autorisation)"),
        'other_non_vat_revenu':fields.float("Others Exempt operations to deduct"),
        'difference':fields.float("Difference"),
        'vat_delivery_onself_revenu': fields.char('Delivery Oneself Revenu to deduct'),
        'taxable_revenu_wout_vat':fields.float("Taxable revenu without tax"),
        'gross_taxable_revenu_wout_vat':fields.float("Taxable revenu without tax"),
        # tva brute
        'normal_rate':fields.float("Rate"),
        "normal_rate_vat_amount":fields.float("Amount"),
        'normal_rate_revenu_amount':fields.function(normal_rate_revenu_amount_calcul,string="Revenu without tax", method=True,type='float', onchange=True),
        'minimal_rate':fields.float("Rate"),
        "minimal_rate_vat_amount":fields.float("Amount"),
        'minimal_rate_revenu_amount':fields.function(minimal_rate_revenu_amount_calcul,string="Revenu without tax", method=True,type='float', onchange=True),
        # regularisation tva 
        'deductible_vat_reserve':fields.float(" PAST DEDUCTIBLE VAT TO REVERSE"),
        'monthly_deductible_vat':fields.float('Monthly deductible vat'),
        'lastest_vat_credit':fields.float(" Vat credit of last month"),
        'deductible_vat_amount':fields.many2one('deductiblevat.statement',"Deductible vat statement",ondelete='set null', track_visibility='onchange',
            select=True, help="Linked Deductible vat statement (optional). Usually created when converting the Deductible vat statement.", required=True),
        'deductible_vat_total':fields.float("TOTAL DEDUCTION"),
        # TOTAL BRUTE
        'vat_gross_total':fields.float("4 - VAT GROSS TOTAL"),
        'vat_to_pay':fields.float("Vat to pay"),
        'credit_vat_toreport':fields.float("Vat credit to report"),
        'credit_vat_torefund':fields.float("Vat credit to refund"),
        
    }
    _defaults = {
    'active': True, 
    'state': 'draft',
    'name':lambda self,cr,uid,context={}: self.pool.get('ir.sequence').get(cr, uid, 'vat.declaration'),
    }
