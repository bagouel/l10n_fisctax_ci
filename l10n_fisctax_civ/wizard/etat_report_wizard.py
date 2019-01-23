

from openerp.osv import  osv, fields, orm 

class deductiblevat_report_file(osv.osv_memory):
	_name = 'deductiblevat.report.file'

	def default_get(self, cr, uid, fields, context=None):
		if context is None:
			context = {}
		res = super(deductiblevat_report_file, self).default_get(cr, uid, fields, context=context)
		res.update({'file_name': context.get('file_name',' Deductible VAT Statement')+'.xlsx'})

		if context.get('file'):
			res.update({'file':context['file']})

		return res
	_columns = {

		'file':fields.binary('File', filters='*.xls'),
		'file_name':fields.char('File Name'),
	}