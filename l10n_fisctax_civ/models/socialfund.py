# -*- coding: utf-8 -*-
from datetime import datetime, timedelta
import time
from openerp.osv import osv, fields, orm
from openerp.tools.translate import _
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT, DEFAULT_SERVER_DATETIME_FORMAT
import openerp.addons.decimal_precision as dp
from openerp import workflow


# Debut de la declaration de la classe cnps dans fiscalité
class socialfund_declaration(osv.osv): 
    _name = 'socialfund.declaration'
    _res_name ='name'

    #Début de la déclaration des fonctions de calculs de CNPS ou SOCIAL FUND
    def on_change_company_elements(self, cr, uid, ids, company_id, context=None):
        values = {}
        if company_id:
            company = self.pool.get('res.company').browse(cr, uid, company_id, context=context)
            values = {
                'address': company.street,
                'phone': company.phone,
                'code_ets': company.code_et,
                'tax_period':company.impot_period,
                'tax_service':company.service,
                'activity_code':company.act_code,
                'employer':company.employeur,
                'company_tax_code':company.vat,
            }
        return {'value': values}
 
    def first_month_wages_total(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.hjoi3231jours1 + record.hjos3231jours1 + record.mi70000mois1 + record.ms70000mois1 + record.ms1647315mois1

        return x

    def first_month_retreat_regime_total(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.hjoi3231jours2 + record.hjos3231jours2 + record.mi70000mois2 + record.ms70000mois2 + record.ms1647315mois2

        return x

    def first_month_benefit_regime_total(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.hjoi3231jours3 + record.hjos3231jours3 + record.mi70000mois3 + record.ms70000mois3 + record.ms1647315mois3

        return x

    def second_month_wages_total(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.hjoi3231jours12 + record.hjos3231jours12 + record.mi70000mois12 + record.ms70000mois12 + record.ms1647315mois12

        return x

    def second_month_retreat_regime_total(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.hjoi3231jours22 + record.hjos3231jours22 + record.mi70000mois22 + record.ms70000mois22 + record.ms1647315mois22

        return x

    def second_month_benefit_regime_total(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.hjoi3231jours32 + record.hjos3231jours32 + record.mi70000mois32 + record.ms70000mois32 + record.ms1647315mois32

        return x

    def third_month_wages_total(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.hjoi3231jours13 + record.hjos3231jours13 + record.mi70000mois13 + record.ms70000mois13 + record.ms1647315mois13

        return x

    def third_month_retreat_regime_total(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.hjoi3231jours23 + record.hjos3231jours23 + record.mi70000mois23 + record.ms70000mois23 + record.ms1647315mois23

        return x

    def third_month_benefit_regime_total(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.hjoi3231jours33 + record.hjos3231jours33 + record.mi70000mois33 + record.ms70000mois33 + record.ms1647315mois33

        return x
    def cumulative_wages_retreat(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.regimetotal + record.regimetotal2 + record.regimetotal3

        return x

    def cumulative_wages_benefit(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.travailtotal + record.travailtotal2 + record.travailtotal3

        return x

    def familly_benefit(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = (record.pf1 * record.pf2)/100

        return x

    def work_accident(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = (record.at1 * record.at2)/100

        return x

    def retreat_regime(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = (record.rr1 * record.rr2)/100

        return x

    def total_to_pay(self, cr, uid, ids, fields, arg, context):
        x = {}
        for record in self.browse(cr, uid, ids):
            x[record.id] = record.pf3 + record.at3 + record.rr3

        return x
    # Fin de la déclaration des fonctions de calcul de CNPS
    _columns = {
         # les champs se trouvant hors des tableaux
        'name': fields.char('Name'),
        'state':fields.selection([('draft','Draft'),('validation','Validation'),('termined','Termined')],'State',required=True,readonly=True),
        'date': fields.date('Date',required=True),
        'periodicity':fields.selection([('monthly','Monthly'),('termly','Termly')],'Periodicity', required=True),
        'tax_period':fields.char("Tax Period", ondelete='cascade'),
        'tax_service':fields.char("Tax Service", ondelete='cascade'),
        'company_id':fields.many2one('res.company','Company',ondelete='set null', track_visibility='onchange',
            select=True, help="Linked company (optional). Usually created when converting the socialfund.", required=True),
        'address':fields.char('Street'),
        'phone':fields.char('Phone'),
        'company_tax_code':fields.char("Tax ID"),
        'code_ets':fields.char('Etablishment Code'),
        'activity_code':fields.char('Activity Code'),
        'employer':fields.char("Employer"),
        
        'totalamount':fields.float("BULLY AMOUNT TOTAL IN THIS PERIOD"),
        # saisie des champs du 1er tableau de la declaration CNPS:
        'month': fields.selection([('january','January'),('february','February'),('march','March'), ('april','April'),('may','May'),('june','June'),('july','July'),('august','August'),('september','September'), ('october','October'),('november','November'),('december','December')],'The First Month of Period',readonly=False, translate=True),
        'month1': fields.selection([('january','January'),('february','February'),('march','March'), ('april','April'),('may','May'),('june','June'),('july','July'),('august','August'),('september','September'), ('october','October'),('november','November'),('december','December')],'The First Month of Period',readonly=False, translate=True),
        'month2': fields.selection([('january','January'),('february','February'),('march','March'), ('april','April'),('may','May'),('june','June'),('july','July'),('august','August'),('september','September'), ('october','October'),('november','November'),('december','December')],'The Second Month of Period',readonly=False, translate=True),
        'month3': fields.selection([('january','January'),('february','February'),('march','March'), ('april','April'),('may','May'),('june','Juin'),('july','July'),('august','August'),('september','September'), ('october','October'),('november','November'),('december','December')],'The Third Month of Period',readonly=False, translate=True),
        'month4': fields.selection([('january','January'),('february','February'),('march','March'), ('april','April'),('may','May'),('june','June'),('july','July'),('august','August'),('september','September'), ('october','October'),('november','November'),('december','December')],'The Second Month of Period',readonly=False, translate=True),
        'month5': fields.selection([('january','January'),('february','February'),('march','March'), ('april','April'),('may','May'),('june','June'),('july','July'),('august','August'),('september','September'), ('october','October'),('november','November'),('december','December')],'The Third Month of Period',readonly=False, translate=True),
        'hjoi3231jours1': fields.integer( ),
        'hjoi3231jours2': fields.float(),
        'hjoi3231jours3': fields.float(),
        'hjos3231jours1': fields.integer(),
        'hjos3231jours2': fields.float(),
        'hjos3231jours3': fields.float(),
        'mi70000mois1': fields.integer(),
        'mi70000mois2': fields.float(),
        'mi70000mois3': fields.float(),
        'ms70000mois1': fields.integer(),
        'ms70000mois2': fields.float(),
        'ms70000mois3': fields.float(),
        'ms1647315mois1':fields.integer(),
        'ms1647315mois2':fields.float(),
        'ms1647315mois3':fields.float(),
        'salaireotal':fields.function(first_month_wages_total, method=True,type='integer', onchange=True),
        'regimetotal':fields.function(first_month_retreat_regime_total, method=True,type='float', onchange=True),
        'travailtotal':fields.function(first_month_benefit_regime_total, method=True,type='float', onchange=True),
        'hjoi3231jours12': fields.integer( ),
        'hjoi3231jours22': fields.float(),
        'hjoi3231jours32': fields.float(),
        'hjos3231jours12': fields.integer(),
        'hjos3231jours22': fields.float(),
        'hjos3231jours32': fields.float(),
        'mi70000mois12': fields.integer(),
        'mi70000mois22': fields.float(),
        'mi70000mois32': fields.float(),
        'ms70000mois12': fields.integer(),
        'ms70000mois22': fields.float(),
        'ms70000mois32': fields.float(),
        'ms1647315mois12':fields.integer(),
        'ms1647315mois22':fields.float(),
        'ms1647315mois32':fields.float(),
        'salaireotal2':fields.function(second_month_wages_total, method=True,type='integer', onchange=True),
        'regimetotal2':fields.function(second_month_retreat_regime_total, method=True,type='float', onchange=True),
        'travailtotal2':fields.function(second_month_benefit_regime_total, method=True,type='float', onchange=True),
        'hjoi3231jours13':fields.integer(),
        'hjoi3231jours23':fields.float(),
        'hjoi3231jours33':fields.float(),
        'hjos3231jours13':fields.integer(),
        'hjos3231jours23':fields.float(),
        'hjos3231jours33':fields.float(),
        'mi70000mois13':fields.integer(),
        'mi70000mois23':fields.float(),
        'mi70000mois33':fields.float(),
        'ms70000mois13':fields.integer(),
        'ms70000mois23':fields.float(),
        'ms70000mois33':fields.float(),
        'ms1647315mois13':fields.integer(),
        'ms1647315mois23':fields.float(),
        'ms1647315mois33':fields.float(),
        'salaireotal3':fields.function(third_month_wages_total, method=True,type='integer', onchange=True),
        'regimetotal3':fields.function(third_month_retreat_regime_total, method=True,type='float', onchange=True),
        'travailtotal3':fields.function(third_month_benefit_regime_total, method=True,type='float', onchange=True),
        'csbsactrr':fields.function(cumulative_wages_retreat,string="Cumulative gross wages contribution for pension plan.", method=True,type='float', onchange=True),
        'csbsctrpf':fields.function(cumulative_wages_benefit,string="Cumulative gross wages contribution for familly benefit and work accident", method=True,type='float', onchange=True),
        # saisie des champs du 2e tableau de la declaration CNPS:
        'pcprr1':fields.integer(),
        'pcprr2':fields.float(),
        'ppqecp1':fields.integer(),
        'ppqecp2':fields.float(),
        'pcprr11':fields.integer(),
        'pcprr12':fields.float(),
        'ppqecp11':fields.integer(),
        'ppqecp12':fields.float(),
        'pcprr21':fields.integer(),
        'pcprr22':fields.float(),
        'ppqecp21':fields.integer(),
        'ppqecp22':fields.float(),


        # saisie des champs du 3e tableau de la declaration CNPS:
        'pf1':fields.function(cumulative_wages_benefit, method=True, type='float', onchange=True),
        'pf2':fields.float(),
        'pf3':fields.function(familly_benefit,string="Contribution for familly prestations", method=True,type='float', onchange=True),
        'at1':fields.function(cumulative_wages_benefit, method=True,type='float', onchange=True),
        'at2':fields.float(),
        'at3':fields.function(work_accident,string="Contribution for work accident", method=True,type='float', onchange=True),
        'rr1':fields.function(cumulative_wages_retreat, method=True,type='float', onchange=True),
        'rr2':fields.float(),
        'rr3':fields.function(retreat_regime,string="Contribution for pension plan", method=True,type='float', onchange=True),
        'tcap':fields.function(total_to_pay,string="TOTAL CONTRIBUTION TO PAY", method=True,type='float', onchange=True),
        
    }
    _defaults = {
    'active': True, 
    'state': 'draft',
    'name':lambda self,cr,uid,context={}: self.pool.get('ir.sequence').get(cr, uid, 'socialfund.declaration'),
    'company_id': lambda s, cr, uid, c: s.pool.get('res.company')._company_default_get(cr, uid, 'socialfund.declaration', context=c),
    }
