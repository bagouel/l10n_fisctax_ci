# -*- coding: utf-8 -*-
from datetime import datetime, timedelta
import time
from openerp.osv import osv, fields, orm
from openerp.tools.translate import _
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT, DEFAULT_SERVER_DATETIME_FORMAT
import openerp.addons.decimal_precision as dp
from openerp import workflow
from datetime import datetime, timedelta
import time
from openerp.osv import osv, fields, orm
from openerp.tools.translate import _
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT, DEFAULT_SERVER_DATETIME_FORMAT
import openerp.addons.decimal_precision as dp
from openerp import workflow


#declaration de la classe contribuable dans fiscalite
class res_company( osv.osv):
    _name = 'res.company'
    _inherit ='res.company'
    _columns =  {
        'sigle':fields.char('Initials'),
        'objet':fields.char('Company Goal'),
        'service':fields.char("Tax Service"),
        'quartier':fields.char("District"),
        'code_et':fields.char("Etablishment Code"),
        'act_code':fields.char("Activity Code"),
        'employeur':fields.char('Employer'),
        'impot_period':fields.char("Tax Period"),
    }
