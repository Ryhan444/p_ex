from odoo import http
from odoo.addons.web.controllers.export import Export, GroupExportXlsxWriter, ExportXlsxWriter, ExcelExport, ExportFormat
import json
import base64
import io
import logging
from odoo.http import request
from PIL import Image
import markupsafe

_logger = logging.getLogger(__name__)

class ExportInherit(Export):
    @http.route('/web/export/formats', type='json', auth="user")
    def formats(self):
        res = super().formats()
        res.append({'tag': 'pdf', 'label': 'PDF'})
        return res

class PdfExport(ExportFormat, http.Controller):
    
    @http.route('/web/export/pdf', type='http', auth="user")
    def web_export_pdf(self, data):
        """Export data ke PDF"""
        try:
            pdf_content = self.base(data)
            return pdf_content
        except Exception as exc:
            _logger.exception("Exception during request handling.")
            payload = json.dumps({
                'code': 200,
                'message': "Odoo Server Error",
                'data': http.serialize_exception(exc)
            })
            raise http.InternalServerError(payload) from exc
    
    @property
    def content_type(self):
        return 'application/pdf'

    @property
    def extension(self):
        return '.pdf'
    
    def _get_pdf_export_html(self, fields, rows, additional_context=None):
        template = "additional_export_pdf_image.non_group_pdf_export_template"
        render_values = {
            'fields': fields,
            'rows': rows,
            'table_start': markupsafe.Markup('<tbody>'),
            'table_end': markupsafe.Markup('''
                </tbody></table>
                <div style="page-break-after: always"></div>
                <table class="o_table table-hover">
            '''),
        }
        return request.env['ir.qweb']._render(template, render_values)
    
    def _get_pdf_group_export_html(self, additional_context=None):
        template = "additional_export_pdf_image.report_group_export_pdf"
        render_values = {
            'data': additional_context,
            'table_start': markupsafe.Markup('<tbody>'),
            'table_end': markupsafe.Markup('''
                </tbody></table>
                <div style="page-break-after: always"></div>
                <table class="o_table table-hover">
            '''),
        }
        return request.env['ir.qweb']._render(template, render_values)
    
    def from_group_data(self, fields, groups):
        base_url = request.env['ir.config_parameter'].sudo().get_param('report.url') or \
                   request.env['ir.config_parameter'].sudo().get_param('web.base.url')
        
        rcontext = {
            'mode': 'print',
            'base_url': base_url,
            'company': request.env.company,
            'fields': fields,
            'groups': self._prepare_group_data(groups),
        }
        
        bodies = [self._get_pdf_group_export_html(additional_context=rcontext)]
        return self._generate_pdf(bodies)
    
    def _prepare_group_data(self, groups, depth=0):
        result = []
        for group_name, group in groups.children.items():
            group_name = group_name[1] if isinstance(group_name, tuple) and len(group_name) > 1 else group_name
            if group._groupby_type[depth] != 'boolean':
                group_name = group_name or _("Undefined")
            child_data = self._prepare_group_data(group, depth + 1)
            result.append({
                'name': group_name,
                'count': group.count,
                'aggregated_values': group.aggregated_values,
                'depth': depth,
                'records': group.data,
            })
            if child_data:
                for child in child_data:
                    result.append({
                        'name': group_name+"/"+child['name'],
                        'count': child['count'],
                        'aggregated_values': child['aggregated_values'],
                        'depth': child['depth'],
                        'records': child['records'],
                    })

        return result
    
    def from_data(self, fields, columns_headers, rows):
        base_url = request.env['ir.config_parameter'].sudo().get_param('report.url') or \
                   request.env['ir.config_parameter'].sudo().get_param('web.base.url')
        
        bodies = [self._get_pdf_export_html(fields=fields, rows=rows, additional_context={'base_url': base_url})]
        return self._generate_pdf(bodies)
    
    def _generate_pdf(self, bodies):
        action_report = request.env['ir.actions.report']
        files_stream = [
            io.BytesIO(action_report._run_wkhtmltopdf(
                bodies,
                footer="",
                landscape=True or self._context.get('force_landscape_printing'),
                specific_paperformat_args={
                    'data-report-margin-top': 10,
                    'data-report-header-spacing': 10,
                    'data-report-margin-bottom': 15,
                }
            ))
        ]
        
        if len(files_stream) > 1:
            result_stream = action_report._merge_pdfs(files_stream)
            result = result_stream.getvalue()
            result_stream.close()
        else:
            result = files_stream[0].read()
        
        for file_stream in files_stream:
            file_stream.close()
        
        return result

def _write_row(self, row, column, data):
    for index, value in enumerate(data):
        current_column = index
        if isinstance(value, bytes):
            try:
                source = base64.b64decode(value)
                image_data = io.BytesIO(source)
                img = Image.open(image_data)
                original_width, original_height = img.size
                cell_width = self.worksheet.default_col_width 
                cell_height = self.worksheet.default_row_height
                x_scale = cell_width * 10 / original_width
                y_scale = cell_height * 6 / original_height
                image_data.seek(0)
                self.worksheet.set_row(row, cell_height * 6)
                self.worksheet.insert_image(row, current_column, "image.png", {
                    "image_data": image_data,
                    "x_scale": x_scale,
                    "y_scale": y_scale
                })
            except Exception as e:
                _logger.error(f"Error processing image: {e}")
        else:
            self.write_cell(row, current_column, value)
    return row + 1, 0

def from_data(self, fields, columns_headers, rows):
    with ExportXlsxWriter(fields, columns_headers, len(rows)) as xlsx_writer:
        for row_index, row in enumerate(rows):
            for cell_index, cell_value in enumerate(row):
                if isinstance(cell_value, bytes):
                    try:
                        source = base64.b64decode(cell_value)
                        image_data = io.BytesIO(source)
                        img = Image.open(image_data)
                        original_width, original_height = img.size
                        cell_width = xlsx_writer.worksheet.default_col_width 
                        cell_height = xlsx_writer.worksheet.default_row_height
                        x_scale = cell_width * 10 / original_width
                        y_scale = cell_height * 6 / original_height
                        image_data.seek(0)
                        xlsx_writer.worksheet.set_row(row_index + 1, cell_height * 6)
                        xlsx_writer.worksheet.insert_image(row_index + 1, cell_index, "image.png", {
                            "image_data": image_data,
                            "x_scale": x_scale,
                            "y_scale": y_scale
                        })
                    except Exception as e:
                        _logger.error(f"Error processing image: {e}")
                else:
                    xlsx_writer.write_cell(row_index + 1, cell_index, cell_value)
    return xlsx_writer.value

GroupExportXlsxWriter._write_row = _write_row
ExcelExport.from_data = from_data
