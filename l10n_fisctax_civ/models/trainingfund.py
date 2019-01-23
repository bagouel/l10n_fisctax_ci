# -*- coding: utf-8 -*-
from datetime import datetime, timedelta
import time
from openerp.osv import osv, fields, orm
from openerp.tools.translate import _
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT, DEFAULT_SERVER_DATETIME_FORMAT
import openerp.addons.decimal_precision as dp
from openerp import workflow



# declaration de la classe fdfp dans fiscalité
class trainingfund_declaration(osv.osv): 
    _name = 'trainingfund.declaration'
    _res_name='name'

    # Déclaration des fonctions de calcul de FDFP ou TRAINING FUND
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

    def monthly_amount_tatf_calcul(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] =(record.remuneration_bttaf * record.rate_taf)/100

        return x

    def monthly_amount_tfcsf_calcul(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] =(record.remuneration_bttf * record.rate_tfcf)/100

        return x

    def payment_todo_calcul(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.monthly_amount_tatf + record.monthly_amount_tfcsf

        return x
    # fin de la déclaration des fonctions de calcul de FDFP ou TRAINING FUND

    _columns = {

        'name': fields.char('Name'),
        'state':fields.selection([('draft','Draft'),('validation','Validation'),('termined','Termined')],'State',required=True,readonly=True),
        'date': fields.date('Date', required=True),
        'periodicity':fields.selection([('monthly','Monthly'),('termly','Termly')],'Periodicity', required=True),
        'tax_period':fields.char("Tax Period"),
        'tax_service':fields.char("Tax Service"),
        'company_id':fields.many2one('res.company','Company Name',ondelete='set null', track_visibility='onchange',
            select=True, help="Linked company (optional). Usually created when converting the trainingfund.", required=True  ),
        'company_tax_code':fields.char('Tax ID'),
        # Effectif des salaires
        'employe_workforce':fields.integer("01- A Workforce"),
        # Determination de la taxe
        'remuneration_bttaf':fields.float(),
        'rate_taf':fields.float(),
        'monthly_amount_tatf':fields.function(monthly_amount_tatf_calcul, string="MONTHLY APPRENTICESHIP TAX AMOUNT", method=True,type='float', onchange=True),
        'remuneration_bttf':fields.float(),
        'rate_tfcf':fields.float(),
        'monthly_amount_tfcsf':fields.function(monthly_amount_tfcsf_calcul, string="MONTHLY AMOUNT (TFC)", method=True,type='float', onchange=True),
        'total_amount_tfcsf':fields.float("2.3- Cumulative TFC from year begin(3)"),
        'regulation_state_sf':fields.char("Regularistion Inside ?"),
        #regularisation annuelle
        'employe_workforce_asf':fields.float("3.1 – Yearly Workforce"),
        'amount_tcfsf':fields.float("3.2 – Monthly amount TFC (3.1 x 1,2 %)"),
        'amount_pay_tcfsf':fields.float("3.3 -Total TFC amount pay during this year (2.3)"),
        'amount_tax_tcfsf':fields.function(monthly_amount_tfcsf_calcul,string="Total tax amount for continuous training", method=True,type='float', onchange=True),
        'engagement_tpsf':fields.float("3.4 - Commitment on FDFP Plan (direct use)"),
        'payment':fields.float("3.5- Payment if |3.2| is under |3.3| + |3.4|"),
        'payment_todo':fields.function(payment_todo_calcul, string="AMOUT TOTAL TO PAY (FDFP)", method=True,type='float', onchange=True),
        
    }
    _defaults = {
    'active': True, 
    'state': 'draft',
    'name':lambda self,cr,uid,context={}: self.pool.get('ir.sequence').get(cr, uid, 'trainingfund.declaration'),
    }
