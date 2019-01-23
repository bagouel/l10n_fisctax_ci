# -*- coding: utf-8 -*-
from datetime import datetime, timedelta
import time
from openerp.osv import osv, fields, orm
from openerp.tools.translate import _
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT, DEFAULT_SERVER_DATETIME_FORMAT
import openerp.addons.decimal_precision as dp
from openerp import workflow


# declaration de la classe etats des taxes deductibles dans fiscalité
class deductiblevat_statement(osv.osv): 
    _name = 'deductiblevat.statement'
    _res_name='name'

    # la fonction de calcul du montant total des elements achetés

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

    def calcul_total_statement(self, cr, uid, ids, field_name, arg, context=None):
        res = {}
        for statement in self.browse(cr, uid, ids, context=context):
            total_statement=0.00
            for line in statement.element_statement_ids:
                total_statement += line.deductible_vat_amount
            res[statement.id] = total_statement
        return res

    _columns = {
        'name': fields.char('Name'),
        'state':fields.selection([('draft','Draft'),('validation','Validation'),('termined','Termined')],'State',required=True,readonly=True),
        'company_id':fields.many2one('res.company',"Company Name",ondelete='set null', track_visibility='onchange',
            select=True, help="Linked company (optional). Usually created when converting the deductiblevat.", required=True),
        'date': fields.date('Date', required=True),
        'periodicity':fields.selection([('monthly','Monthly'),('termly','Termly')],'Periodicity', required=True),
        'tax_period':fields.char("Tax Period"),
        'tax_service':fields.char("Tax Service",),
        'company_tax_code':fields.char( 'Tax ID',),
        'total_statement':fields.function(calcul_total_statement, string='Total', method=True, type='float', onchange=True ),

        # ajout des elements d'etat
        'element_statement_ids': fields.one2many('deductiblevat.statement.lines', 'element_id', 'Order Lines'),
    }
    _defaults = {
    'active': True, 
    'state': 'draft',
    'name':lambda self,cr,uid,context={}: self.pool.get('ir.sequence').get(cr, uid, 'deductiblevat.statement'),
    }