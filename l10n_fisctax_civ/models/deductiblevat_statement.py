# -*- coding: utf-8 -*-
from datetime import datetime, timedelta
import time
from openerp.osv import osv, fields, orm
from openerp.tools.translate import _
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT, DEFAULT_SERVER_DATETIME_FORMAT
import openerp.addons.decimal_precision as dp
from openerp import workflow



# declaration de la classe elements etats des taxes deductibles dans fiscalité

class deductiblevat_statement_lines(osv.osv): 
    _name = 'deductiblevat.statement.lines'

    # declaration des fonctions de calcul des elements des taxes déductibles

    def func_vat_amount(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = (record.amount * record.annual_tax)/100

        return x


    def func_deductible_vat_amount(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.vat_amount - (record.vat_amount * record.deduction_rate)

        return x

    _columns = {
        'annual_tax':fields.float('Annual tax (In %)',track_visibility='onchange'),
        'element_id':fields.many2one('deductiblevat.statement','element_statement', 'Order Reference', required=True, ondelete='cascade', domain=[('sale_ok', '=', True)], readonly=True),
        'date':fields.date("Invoice or document date"),
        'supplier':fields.char("Supplier Name"),
        'supplier_taxt_code':fields.char("Supplier Tax ID"),
        'document_ref':fields.char("Invoice or Document Reference"),
        'goods_definition':fields.char("Definition of Goods or Services (subject to deduction)"),
        'amount':fields.float("Amount Total", track_visibility='onchange'),
        'vat_amount':fields.function(func_vat_amount,string="VAT Amount",track_visibility='onchange', method=True, type='float', onchange=True),
        'deduction_rate':fields.float("Deduction Rate",track_visibility='onchange'),
        'deductible_vat_amount':fields.function(func_deductible_vat_amount,string="Deductible VAT Amount", method=True,type='float'), 
    }
