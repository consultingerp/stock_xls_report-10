from odoo.addons.report_xlsx.report.report_xlsx import ReportXlsx
from odoo import fields, models,api
from collections import OrderedDict
import datetime
from datetime import date, datetime


class StockReportXls(ReportXlsx):
    @api.multi
    def get_lines(self, date_from, date_to, product_categ, product, cat, warehouse,locations,check):
         
        lines = []
        imf=[]
        if product:
            product_ids = self.env['product.product'].search([('categ_id','=',cat.id),('id','in',product)])
        else:
            product_ids = self.env['product.product'].search([('categ_id','=',cat.id)])
        
        array = []
        for lo in locations:
            array.append(lo.id)
        for product in product_ids:
            cost = product.standard_price
            sale_price = product.list_price
            product_source =  self.env['stock.move'].search([('date', '>=',date_from),('date', '<=',date_to),('product_id', '=',product.id),('location_id','in',array),('state','=','done')])
            product_dest =  self.env['stock.move'].search([('date', '>=',date_from),('date', '<=',date_to),('product_id', '=',product.id),('location_dest_id','in',array),('state','=','done')])
            production = 0.0
            itp = 0.0
            sale = 0.0
            inventory_loss = 0.0
            inventory_loss1 = 0.0
            transfer = 0.0
            purchase =0.0 
            transfer_two = 0.0
            purchase_return = 0.0
            sale_return = 0.0
            array_location = []
            
            for prod in product_source:
                if prod.location_id.usage == 'internal' and prod.location_dest_id.usage =='production':
                    itp += prod.product_uom_qty
                if prod.location_id.usage == 'internal' and prod.location_dest_id.usage =='customer':
                    sale += prod.product_uom_qty
#                     or prod.location_dest_id.usage =='view' or prod.location_dest_id.usage =='procurement' or prod.location_dest_id.usage =='transit'
                if prod.location_id.usage == 'internal' and prod.location_dest_id.usage =='inventory':
                    inventory_loss += prod.product_uom_qty
                if prod.location_id.usage == 'internal' and prod.location_dest_id.usage =='internal':
                    transfer = transfer+prod.product_uom_qty
                if prod.location_id.usage == 'internal' and prod.location_dest_id.usage =='supplier':
                    purchase_return += prod.product_uom_qty
            for dest in product_dest:
                if dest.location_id.usage == 'supplier' and dest.location_dest_id.usage == 'internal':
                    purchase += dest.product_uom_qty
                if dest.location_id.usage == 'production' and dest.location_dest_id.usage == 'internal':
                    production += dest.product_uom_qty 
                if dest.location_id.usage == 'internal' and dest.location_dest_id.usage == 'internal':
                    transfer_two += dest.product_uom_qty
                if dest.location_id.usage == 'customer' and dest.location_dest_id.usage == 'internal':
                    sale_return += dest.product_uom_qty
                if dest.location_id.usage == 'inventory' and dest.location_dest_id.usage =='internal':
                    inventory_loss1 += dest.product_uom_qty
            
            product_source1 =  self.env['stock.move'].search([('date', '<',date_from),('product_id', '=',product.id),('location_id','in',array),('state','=','done')])
            product_dest1 =  self.env['stock.move'].search([('date', '<',date_from),('product_id', '=',product.id),('location_dest_id','in',array),('state','=','done')])
            ob_production = 0.0
            ob_itp = 0.0
            ob_sale = 0.0
            ob_inventory_loss = 0.0
            ob_inventory_loss1 = 0.0
            ob_transfer = 0.0
            ob_purchase =0.0 
            ob_transfer_two = 0.0
            opening_bal= 0.0
            ob_purchase_return = 0.0
            ob_sale_return = 0.0
            closing_bal= 0.0
            
            for prod in product_source1:
                if prod.location_id.usage == 'internal' and prod.location_dest_id.usage =='production':
                    ob_itp += prod.product_uom_qty
                if prod.location_id.usage == 'internal' and prod.location_dest_id.usage =='customer':
                    ob_sale += prod.product_uom_qty
#                     or prod.location_dest_id.usage =='view' or prod.location_dest_id.usage =='procurement' or prod.location_dest_id.usage =='transit'
                if prod.location_id.usage == 'internal' and prod.location_dest_id.usage =='inventory':
                    ob_inventory_loss += prod.product_uom_qty
                if prod.location_id.usage == 'internal' and prod.location_dest_id.usage =='internal':
                        ob_transfer += prod.product_uom_qty
                if prod.location_id.usage == 'internal' and prod.location_dest_id.usage =='supplier':
                    ob_purchase_return += prod.product_uom_qty
            for dest in product_dest1:
                if dest.location_id.usage == 'supplier' and dest.location_dest_id.usage == 'internal':
                    ob_purchase += dest.product_uom_qty
                if dest.location_id.usage == 'production' and dest.location_dest_id.usage == 'internal':
                    ob_production += dest.product_uom_qty 
                if dest.location_id.usage == 'internal' and dest.location_dest_id.usage == 'internal':
                    ob_transfer_two += dest.product_uom_qty
                if dest.location_id.usage == 'customer' and dest.location_dest_id.usage == 'internal':
                    ob_sale_return += dest.product_uom_qty
                if dest.location_id.usage == 'inventory' and dest.location_dest_id.usage =='internal':
                    ob_inventory_loss1 += dest.product_uom_qty
                
            opening_bal= ob_purchase - ob_sale + ob_inventory_loss1 + ob_sale_return - ob_purchase_return + ob_production
            closing_bal = opening_bal + purchase + production + transfer + inventory_loss1 - inventory_loss - itp - sale
#                 if 'MO' in prod.name:
            if check:
                vals = {
                    'code': product.default_code or ' ',
                    'name': product.name + ' ' + str(product.attribute_value_ids.name or ' '),
                    'production': production*cost or 0,
                    'purchase': purchase*cost or 0,
                    'sale': sale*sale_price,
                    'transfer': (transfer_two - transfer)*cost,
                    'inventory_loss': (inventory_loss1 - inventory_loss)*cost,
                    'itp': itp*cost,
                    'opening_bal': opening_bal*cost,
                    'closing_bal': closing_bal*cost
                }
            else:
                vals = {
                        'code': product.default_code or ' ',
                        'name': product.name + ' ' + str(product.attribute_value_ids.name or ' '),
                        'production':production or 0,
                        'purchase': purchase or 0,
                        'sale': sale,
                        'transfer': transfer_two - transfer,
                        'inventory_loss': inventory_loss1 - inventory_loss,
                        'itp': itp,
                        'opening_bal': opening_bal,
                        'closing_bal':closing_bal
                        }
            lines.append(vals)
        return lines

    def generate_xlsx_report(self, workbook, data, lines):
        sheet = workbook.add_worksheet()
#         report_name = data['form']['report_type']
        
        format1 = workbook.add_format({'font_size': 14, 'bottom': True, 'right': True, 'left': True, 'top': True, 'align': 'center', 'bold': True})
        format11 = workbook.add_format({'font_size': 14, 'align': 'center', 'bold': True,})
#         format123 = workbook.set_column('A:A', 100)
        period_format= workbook.add_format({'font_size': 11, 'align': 'center', 'bold': True})

        format12 = workbook.add_format({'font_size': 11, 'align': 'center', 'bold': True,'right': True, 'left': True,'bottom': True, 'top': True})
        format21 = workbook.add_format({'font_size': 10, 'bold': True, 'align': 'right', 'right': True, 'left': True,'bottom': True, 'top': True})
        format21.set_num_format('#,##0.00')
        qty_format = workbook.add_format({'font_size': 10, 'align': 'right', 'right': True, 'left': True,'bottom': True, 'top': True})
        qty_format.set_num_format('#,##0.000')
        Pname_format = workbook.add_format({'font_size': 10, 'align': 'left', 'right': True, 'left': True,'bottom': True, 'top': True})
        format_center = workbook.add_format({'font_size': 10, 'align': 'center', 'right': True, 'left': True,'bottom': True, 'top': True})
        subtotal_format = workbook.add_format({'font_size': 10, 'bold': True, 'align': 'right', 'right': True, 'left': True,'bottom': True, 'top': True})
        subtotal_format.set_num_format('#,##0.000')
        font_size_8 = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 8})
        red_mark = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 8,
                                        'bg_color': 'red'})
        justify = workbook.add_format({'bottom': True, 'top': True, 'right': True, 'left': True, 'font_size': 12})
#         style = workbook.add_format('align: wrap yes; borders: top thin, bottom thin, left thin, right thin;')
#         style.num_format_str = '#,##0.00'
#         format3.set_align('center')
        font_size_8.set_align('center')
        justify.set_align('justify')
        format1.set_align('center') 
        red_mark.set_align('center')
                
        date_from = datetime.strptime(data['form']['date_from'], '%Y-%m-%d').strftime('%d/%m/%y')
        date_to = datetime.strptime(data['form']['date_to'], '%Y-%m-%d').strftime('%d/%m/%y')
#         if report_name == 'grand_production_summary':
        sheet.merge_range(0, 0, 0, 9, 'Brand Wise Summary of Stock Movement', format11)
        sheet.merge_range(1, 0, 1, 9, 'Period from: ' + (date_from) +  ' to ' + (date_to), period_format)
     
        # report start
        product_row = 5
        cat_row = 2
        warehouse = data['form']['warehouse']
        category = self.env['product.category'].search([])
        product_categ = data['form']['product_categ']
        if product_categ:
            category = self.env['product.category'].search([('id','in',product_categ)])
        else:
            category = self.env['product.category'].search([])
        if warehouse:
            warehouses = self.env['stock.warehouse'].search([('id','=',warehouse)])
               
#         else:
#             locations = self.env['stock.location'].search([])
#             warehouses = self.env['stock.warehouse'].search([])
#         if report_name == 'grand_production_summary':
        for ware in warehouses:
            
            if data['form']['location']:
                locations = self.env['stock.location'].search([('id','in',data['form']['location'])])
            else:
                locations = self.env['stock.location'].search([('Wr_id','=',warehouse)])
#             array1 = []
#             for lo in locations:
#                 array1.append(lo.id)
#             product_data=  self.env['mrp.production'].search([('create_date', '>=',data['form']['date_from']),('create_date', '<=',data['form']['date_to'])])
#             if product_data:
            sheet.merge_range(product_row-3, 0, product_row-3, 2,ware.name, format12)
        
            for cat in category:
                
                get_lines = self.get_lines(data['form']['date_from'],data['form']['date_to'],category,data['form']['product'],cat,ware,locations,data['form']['check'])
        #        
                if get_lines:  
                    total_ob = 0.0
                    total_pu=0.0
                    total_p=0.0
                    total_t=0.0
                    total_adj=0.0
                    total_itp=0.0
                    total_s=0.0
                    total_cb=0.0
                    sheet.write(product_row-2, 0, 'Category', format12)
                    sheet.merge_range(product_row-2, 1,product_row-2, 2,cat.name, format12)
                    
                    sheet.write(product_row-1, 0,'Code', format12)
                    sheet.write(product_row-1, 1,'Name', format12)
                    sheet.write(product_row-1, 2,'Opening Balance', format12)
                    sheet.write(product_row-1, 3,'Purchase', format12)
                    sheet.write(product_row-1, 4,'Production', format12)
                    sheet.write(product_row-1, 5,'Transfers', format12)
                    sheet.write(product_row-1, 6,'Adjustments (+/-)', format12)
                    sheet.write(product_row-1, 7,'Issue to Production', format12)
                    sheet.write(product_row-1, 8,'Sale', format12)
                    sheet.write(product_row-1, 9,'Closing Balance', format12)
                
                   
                    
#                     temp_code='None'
#                     product_code = []
                    for line in get_lines:
                        
                        
#                         if temp_code != line['code']:
#                             for pro_line in get_lines:
#                                 if pro_line['code'] == line['code']:
#                                     total1+=pro_line['production']
                        sheet.write(product_row, 0, line['code'], format_center)
                        sheet.write(product_row, 1, line['name'], Pname_format)
                        sheet.write(product_row, 2, line['opening_bal'], qty_format)
                        sheet.write(product_row, 3, line['purchase'], qty_format)
                        sheet.write(product_row, 4, line['production'], qty_format)
                        sheet.write(product_row, 5, line['transfer'], qty_format)
                        sheet.write(product_row, 6, line['inventory_loss'], qty_format)
                        sheet.write(product_row, 7, line['itp'], qty_format)
                        sheet.write(product_row, 8, line['sale'], qty_format)
                        sheet.write(product_row, 9, line['closing_bal'], qty_format)

                        product_row +=1
                             
#                         total1=0    
#                     total+=line['production']
#                     temp_code= line['code']
                        total_ob +=line['opening_bal']
                        total_pu += line['purchase']
                        total_p += line['production']
                        total_t += line['transfer']
                        total_adj += line['inventory_loss']
                        total_itp += line['itp']
                        total_s += line['sale']
                        total_cb += line['closing_bal']
                        
                    sheet.write(product_row, 1,'Sub Total', format21)
                    sheet.write(product_row, 2, total_ob, subtotal_format)
                    sheet.write(product_row, 3, total_pu, subtotal_format)
                    sheet.write(product_row, 4, total_p, subtotal_format)
                    sheet.write(product_row, 5, total_t, subtotal_format)
                    sheet.write(product_row, 6, total_adj, subtotal_format)
                    sheet.write(product_row, 7, total_itp, subtotal_format)
                    sheet.write(product_row, 8, total_s, subtotal_format)
                    sheet.write(product_row, 9, total_cb, subtotal_format)

                    product_row+=4

StockReportXls('report.export_stockinfo_xls.stock_report_xls.xlsx', 'product.product')
