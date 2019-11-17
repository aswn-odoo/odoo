
import base64
from datetime import datetime
from datetime import *
from io import BytesIO

import xlsxwriter
from odoo import fields, models, api, _


class payroll_report_excel(models.TransientModel):
    _name = 'account.report.excel'

    name = fields.Char('File Name', size=256, readonly=True)
    file_download = fields.Binary('Download payroll', readonly=True)

class AccountingReport(models.TransientModel):
    _inherit = "account.report.general.ledger"

    def _get_account_move_entry(self, accounts, init_balance, sortby, display_account):
        cr = self.env.cr
        MoveLine = self.env['account.move.line']
        move_lines = {x: [] for x in accounts.ids}
        print(self.env.context)
        # Prepare initial sql query and Get the initial move lines
        if init_balance:
            init_tables, init_where_clause, init_where_params = MoveLine.with_context(date_from=self.env.context.get('date_from'), date_to=False, initial_bal=True)._query_get()
            init_wheres = [""]
            if init_where_clause.strip():
                init_wheres.append(init_where_clause.strip())
            init_filters = " AND ".join(init_wheres)
            filters = init_filters.replace('account_move_line__move_id', 'm').replace('account_move_line', 'l')
            sql = ("""SELECT 0 AS lid, l.account_id AS account_id, '' AS ldate, '' AS lcode, 0.0 AS amount_currency,\
                 '' AS lref, 'Initial Balance' AS lname,\
                 COALESCE(SUM(l.debit),0.0) AS debit,\
                 COALESCE(SUM(l.credit),0.0) AS credit,\
                 COALESCE(SUM(l.debit),0) - COALESCE(SUM(l.credit), 0) as balance, '' AS lpartner_id,\
                '' AS move_name, '' AS mmove_id, '' AS currency_code,\
                NULL AS currency_id,\
                '' AS invoice_id, '' AS invoice_type, '' AS invoice_number,\
                '' AS partner_name\
                FROM account_move_line l\
                LEFT JOIN account_move m ON (l.move_id=m.id)\
                LEFT JOIN res_currency c ON (l.currency_id=c.id)\
                LEFT JOIN res_partner p ON (l.partner_id=p.id)\
                LEFT JOIN account_invoice i ON (m.id =i.move_id)\
                JOIN account_journal j ON (l.journal_id=j.id)\
                WHERE l.account_id IN %s""" + filters + ' GROUP BY l.account_id')
            params = (tuple(accounts.ids),) + tuple(init_where_params)
            cr.execute(sql, params)
            for row in cr.dictfetchall():
                move_lines[row.pop('account_id')].append(row)

        sql_sort = 'l.date, l.move_id'
        if sortby == 'sort_journal_partner':
            sql_sort = 'j.code, p.name, l.move_id'

        # Prepare sql query base on selected parameters from wizard
        tables, where_clause, where_params = MoveLine._query_get()
        wheres = [""]
        if where_clause.strip():
            wheres.append(where_clause.strip())
        filters = " AND ".join(wheres)
        filters = filters.replace('account_move_line__move_id', 'm').replace('account_move_line', 'l')

        # Get move lines base on sql query and Calculate the total balance of move lines
        sql = ('''SELECT l.id AS lid, l.account_id AS account_id, l.date AS ldate, j.code AS lcode, l.currency_id, \
                l.amount_currency, l.ref AS lref, l.name AS lname, COALESCE(l.debit,0) AS debit,\
                COALESCE(l.credit,0) AS credit,\
                COALESCE(SUM(l.debit),0) - COALESCE(SUM(l.credit), 0) \
                AS balance,\
            m.name AS move_name, c.symbol AS currency_code, p.name AS partner_name\
            FROM account_move_line l JOIN account_move m ON (l.move_id=m.id)\
            LEFT JOIN res_currency c ON (l.currency_id=c.id) LEFT JOIN res_partner p ON (l.partner_id=p.id)\
            JOIN account_journal j ON (l.journal_id=j.id) JOIN account_account acc ON (l.account_id = acc.id) \
            WHERE l.account_id IN %s ''' + filters + ''' GROUP BY l.id, l.account_id, l.date, j.code, l.currency_id,\
                 l.amount_currency, l.ref, l.name, m.name, c.symbol, p.name ORDER BY ''' + sql_sort)
        params = (tuple(accounts.ids),) + tuple(where_params)
        cr.execute(sql, params)

        for row in cr.dictfetchall():
            balance = 0
            for line in move_lines.get(row['account_id']):
                balance += line['debit'] - line['credit']
            row['balance'] += balance
            move_lines[row.pop('account_id')].append(row)

        # Calculate the debit, credit and balance for Accounts
        account_res = []
        for account in accounts:
            currency = account.currency_id and account.currency_id or account.company_id.currency_id
            res = dict((fn, 0.0) for fn in ['credit', 'debit', 'balance'])
            res['code'] = account.code
            res['name'] = account.name
            res['move_lines'] = move_lines[account.id]
            for line in res.get('move_lines'):
                res['debit'] += line['debit']
                res['credit'] += line['credit']
                res['balance'] = line['balance']
            if display_account == 'all':
                account_res.append(res)
            if display_account == 'movement' and res.get('move_lines'):
                account_res.append(res)
            if display_account == 'not_zero' and not currency.is_zero(res['balance']):
                account_res.append(res)

        return account_res


    @api.multi
    def print_excel_report(self):
        self.ensure_one()
        data = {}
        data['ids'] = self.env.context.get('active_ids', [])
        data['model'] = self.env.context.get('active_model', 'ir.ui.menu')
        data['form'] = self.read(['date_from', 'date_to', 'journal_ids', 'target_move', 'company_id', 'init_balance', 'sortby', 'display_account'])[0]
        used_context = self._build_contexts(data)
        data['form']['used_context'] = dict(used_context, lang=self.env.context.get('lang') or 'en_US')
        if data['form'].get('journal_ids', False):
            codes = [journal.code for journal in self.env['account.journal'].search([('id', 'in', data['form'].get('journal_ids'))])]

        accounts = self.env['account.account'].search([])
        accounts_res = self.with_context(data['form'].get('used_context'))._get_account_move_entry(accounts, data['form'].get('init_balance'), data['form'].get('sortby'), data['form'].get('display_account'))





        from pprint import pprint
        pprint(accounts_res)
        file_name = _('General Ledger report.xlsx')
        fp = BytesIO()
        workbook = xlsxwriter.Workbook(fp)
        heading_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True, 'size': 14})
        cell_text_format_n = workbook.add_format({'align': 'center', 'bold': True, 'size': 9})
        cell_text_value_n = workbook.add_format({'align': 'center', 'size': 9})
        cell_text_format = workbook.add_format({'align': 'left', 'bold': True, 'size': 12})
        table_cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'size': 11})
        cell_text_format_new = workbook.add_format({'align': 'left', 'size': 11})
        header_format = workbook.add_format({'align': 'center', 'bold': True, 'size': 12})
        header_format.set_border()
        table_cell_format.set_border()
        cell_text_format.set_border()
        cell_text_format_new.set_border()
        worksheet = workbook.add_worksheet('General Ledger report.xlsx')
        date_2 = datetime.strftime(self.date_to, '%d-%m-%Y')
        date_1= datetime.strftime(self.date_from, '%d-%m-%Y')
        sort_by =  'Date' if data['form'].get('sortby') == 'sort_date' else 'Journal & Partner'
        state =  'All Entries' if data['form'].get('target_move') == 'all' else 'All Posted Entries'
        disply_account = data['form'].get('display_account').replace('_', ' ').capitalize()
        worksheet.merge_range('A1:I2', '%s : %s' % (self.company_id.name, 'General Ledger'), heading_format)
        row = 2
        worksheet.write(row, 0, 'Journals', cell_text_format_n)
        worksheet.merge_range('B3:D3',','.join(codes), cell_text_value_n)
        worksheet.write(row+1, 0, 'Sorted By', cell_text_format_n)
        worksheet.write(row+1, 1, sort_by or '', cell_text_value_n)
        worksheet.write(row, 5, 'Display Account', cell_text_format_n)
        worksheet.write(row, 6, disply_account or '', cell_text_value_n)
        worksheet.write(row+1, 5, 'Target Move', cell_text_format_n)
        worksheet.write(row+1, 6, state or '', cell_text_value_n)
        worksheet.write(row, 7, 'Date From', cell_text_format_n)
        worksheet.write(row, 8, date_1 or '',cell_text_value_n)
        worksheet.write(row+1, 7, 'Date To', cell_text_format_n)
        worksheet.write(row+1, 8, date_2 or '', cell_text_value_n)
        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 10)
        worksheet.set_column('C:C', 20)
        worksheet.set_column('D:D', 30)
        worksheet.set_column('E:E', 30)
        worksheet.set_column('F:F', 40)
        worksheet.set_column('G:G', 10)
        worksheet.set_column('H:H', 10)
        worksheet.set_column('I:I', 10)
        row +=4
        col= 0
        table_headers = ['Date', 'JNRL', 'Partner', 'Ref', 'Move', 'Entry Label', 'Debit', 'Credit','Balance']
        for header in table_headers:
            worksheet.write(row, col, header, header_format)
            col +=1
        row +=1
        for rec in accounts_res:
            worksheet.merge_range(row,0,row,5,rec.get('name'), cell_text_format)
            worksheet.write(row,6,rec.get('debit'), cell_text_format_new)
            worksheet.write(row,7,rec.get('credit'), cell_text_format_new)
            worksheet.write(row,8,rec.get('balance'), cell_text_format_new)
            row +=1
            for line in rec.get('move_lines'):
                worksheet.write(row,0,str(line.get('ldate')), table_cell_format)
                worksheet.write(row,1,line.get('lcode'), table_cell_format)
                worksheet.write(row,2,line.get('partner_name'), table_cell_format)
                worksheet.write(row,3,line.get('lref'), table_cell_format)
                worksheet.write(row,4,line.get('move_name'), table_cell_format)
                worksheet.write(row,5,line.get('lname'), cell_text_format_new)
                worksheet.write(row,6,line.get('debit'), cell_text_format_new)
                worksheet.write(row,7,line.get('credit'), cell_text_format_new)
                worksheet.write(row,8,line.get('balance'), cell_text_format_new)
                row +=1
        workbook.close()
        file_download = base64.b64encode(fp.getvalue())
        fp.close()
        excel_report = self.env['account.report.excel'].create({'name' : file_name,
                                                'file_download' : file_download})
        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/?model=account.report.excel&id={}&field=file_download&filename_field=name&download=true'.format(excel_report.id),
        }

class AccountPartnerLedger(models.TransientModel):
    _inherit = "account.report.partner.ledger"

    def _lines(self, data, partner):
        full_account = []
        currency = self.env['res.currency']
        query_get_data = self.env['account.move.line'].with_context(data['form'].get('used_context', {}))._query_get()
        reconcile_clause = "" if data['form']['reconciled'] else ' AND "account_move_line".full_reconcile_id IS NULL '
        params = [partner.id, tuple(data['computed']['move_state']), tuple(data['computed']['account_ids'])] + query_get_data[2]
        query = """
            SELECT "account_move_line".id, "account_move_line".date, j.code, acc.code as a_code, acc.name as a_name, "account_move_line".ref, m.name as move_name, "account_move_line".name, "account_move_line".debit, "account_move_line".credit, "account_move_line".amount_currency,"account_move_line".currency_id, c.symbol AS currency_code
            FROM """ + query_get_data[0] + """
            LEFT JOIN account_journal j ON ("account_move_line".journal_id = j.id)
            LEFT JOIN account_account acc ON ("account_move_line".account_id = acc.id)
            LEFT JOIN res_currency c ON ("account_move_line".currency_id=c.id)
            LEFT JOIN account_move m ON (m.id="account_move_line".move_id)
            WHERE "account_move_line".partner_id = %s
                AND m.state IN %s
                AND "account_move_line".account_id IN %s AND """ + query_get_data[1] + reconcile_clause + """
                ORDER BY "account_move_line".date"""
        self.env.cr.execute(query, tuple(params))
        res = self.env.cr.dictfetchall()
        sum = 0.0
        lang_code = self.env.context.get('lang') or 'en_US'
        lang = self.env['res.lang']
        lang_id = lang._lang_get(lang_code)
        date_format = lang_id.date_format
        for r in res:
            r['date'] = r['date']
            r['displayed_name'] = '-'.join(
                r[field_name] for field_name in ('move_name', 'ref', 'name')
                if r[field_name] not in (None, '', '/')
            )
            sum += r['debit'] - r['credit']
            r['progress'] = sum
            r['currency_id'] = currency.browse(r.get('currency_id'))
            full_account.append(r)
        return full_account

    def _sum_partner(self, data, partner, field):
        if field not in ['debit', 'credit', 'debit-credit']:
            return
        result = 0.0
        query_get_data = self.env['account.move.line'].with_context(data['form'].get('used_context', {}))._query_get()
        reconcile_clause = "" if data['form']['reconciled'] else ' AND "account_move_line".full_reconcile_id IS NULL '

        params = [partner.id, tuple(data['computed']['move_state']), tuple(data['computed']['account_ids'])] + query_get_data[2]
        query = """SELECT sum(""" + field + """)
                FROM """ + query_get_data[0] + """, account_move AS m
                WHERE "account_move_line".partner_id = %s
                    AND m.id = "account_move_line".move_id
                    AND m.state IN %s
                    AND account_id IN %s
                    AND """ + query_get_data[1] + reconcile_clause
        self.env.cr.execute(query, tuple(params))

        contemp = self.env.cr.fetchone()
        if contemp is not None:
            result = contemp[0] or 0.0
        return result

    @api.multi
    def print_excel_report(self):
        self.ensure_one()
        data = {}
        data['ids'] = self.env.context.get('active_ids', [])
        data['model'] = self.env.context.get('active_model', 'ir.ui.menu')
        data['form'] = self.read(['date_from', 'date_to', 'journal_ids', 'target_move', 'company_id', 'reconciled', 'amount_currency'])[0]
        used_context = self._build_contexts(data)
        data['form']['used_context'] = dict(used_context, lang=self.env.context.get('lang') or 'en_US')
        from pprint import pprint
        pprint(data)
        data['computed'] = {}
        obj_partner = self.env['res.partner']
        query_get_data = self.env['account.move.line'].with_context(data['form'].get('used_context', {}))._query_get()
        data['computed']['move_state'] = ['draft', 'posted']
        if data['form'].get('target_move', 'all') == 'posted':
            data['computed']['move_state'] = ['posted']
        result_selection = data['form'].get('result_selection', 'customer')
        if result_selection == 'supplier':
            data['computed']['ACCOUNT_TYPE'] = ['payable']
        elif result_selection == 'customer':
            data['computed']['ACCOUNT_TYPE'] = ['receivable']
        else:
            data['computed']['ACCOUNT_TYPE'] = ['payable', 'receivable']

        self.env.cr.execute("""
            SELECT a.id
            FROM account_account a
            WHERE a.internal_type IN %s
            AND NOT a.deprecated""", (tuple(data['computed']['ACCOUNT_TYPE']),))
        data['computed']['account_ids'] = [a for (a,) in self.env.cr.fetchall()]
        params = [tuple(data['computed']['move_state']), tuple(data['computed']['account_ids'])] + query_get_data[2]
        reconcile_clause = "" if data['form']['reconciled'] else ' AND "account_move_line".full_reconcile_id IS NULL '
        query = """
            SELECT DISTINCT "account_move_line".partner_id
            FROM """ + query_get_data[0] + """, account_account AS account, account_move AS am
            WHERE "account_move_line".partner_id IS NOT NULL
                AND "account_move_line".account_id = account.id
                AND am.id = "account_move_line".move_id
                AND am.state IN %s
                AND "account_move_line".account_id IN %s
                AND NOT account.deprecated
                AND """ + query_get_data[1] + reconcile_clause
        self.env.cr.execute(query, tuple(params))
        partner_ids = [res['partner_id'] for res in self.env.cr.dictfetchall()]
        partners = obj_partner.browse(partner_ids)
        partners = sorted(partners, key=lambda x: (x.ref or '', x.name or ''))
        fp = BytesIO()
        workbook = xlsxwriter.Workbook(fp)

        heading_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True, 'size': 14})
        header_format = workbook.add_format({'align': 'center', 'bold': True, 'size': 12})
        heading_format.set_border()
        worksheet = workbook.add_worksheet('Partner Ledger report.xlsx')
        date_2 = datetime.strftime(self.date_to, '%d-%m-%Y')
        date_1= datetime.strftime(self.date_from, '%d-%m-%Y')
        state =  'All Entries' if data['form'].get('target_move') == 'all' else 'All Posted Entries'
        worksheet.merge_range('A1:H2', '%s : %s' % (self.company_id.name, 'General Ledger'), heading_format)
        company = self.company_id.name
        row = 2
        worksheet.write(row, 0, 'Company', header_format)
        worksheet.merge_range(row, 1, row, 2, company or '')
        worksheet.write(row, 3, 'Date From', header_format)
        worksheet.write(row, 4, date_1 or '')
        worksheet.write(row+1, 3, 'Date To', header_format)
        worksheet.write(row+1, 4, date_2 or '')
        worksheet.write(row, 6, 'Target Move', header_format)
        worksheet.write(row, 7, date_1 or '')
        row += 2
        col = 0
        worksheet.set_column('D:D', 30)
        worksheet.set_column('G:G', 20)
        worksheet.set_column('E:E', 20)
        worksheet.set_column('H:H', 20)
        worksheet.set_column('A:A', 30)
        worksheet.write(row,0,'Date', header_format)
        worksheet.write(row,1,'JRNL', header_format)
        worksheet.write(row,2,'Account', header_format)
        worksheet.write(row,3,'Ref', header_format)
        worksheet.write(row,4,'Debit', header_format)
        worksheet.write(row,5,'Credit', header_format)
        worksheet.write(row,6,'Balance', header_format)
        worksheet.write(row,7,'Currency', header_format)
        row +=1
        for partner in partners:
            worksheet.merge_range(row,0,row,3,partner.ref if partner.ref else ' '+'-'+partner.name)
            worksheet.write(row,4,self._sum_partner(data,partner, 'debit'))
            worksheet.write(row,5,self._sum_partner(data,partner, 'credit'))
            worksheet.write(row,6,self._sum_partner(data,partner, 'debit-credit'))
            row +=1
            for line in self._lines(data, partner):
                worksheet.write(row, 0, str(line['date']))
                worksheet.write(row, 1,line['code'])
                worksheet.write(row, 2,line['a_code'])
                worksheet.write(row, 3,line['displayed_name'])
                worksheet.write(row, 4,line['debit'])
                worksheet.write(row, 5,line['credit'])
                worksheet.write(row, 6,line['progress'])
                if line['currency_id']:
                    worksheet.write(row, 7,line['amount_currency'])
                row +=1


        workbook.close()
        report_id = self.env['account.report.excel'].create({
            'file_download': base64.encodestring(fp.getvalue()),
            'name': 'partner_ledger_xls''.xlsx'
        })
        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/?model=account.report.excel&id={}&field=file_download&filename_field=name&download=true'.format(report_id.id),
        }

class AccountBalanceReport(models.TransientModel):
    _inherit = "account.balance.report"

    def _get_accounts(self, accounts, display_account):
        """ compute the balance, debit and credit for the provided accounts
            :Arguments:
                `accounts`: list of accounts record,
                `display_account`: it's used to display either all accounts or those accounts which balance is > 0
            :Returns a list of dictionary of Accounts with following key and value
                `name`: Account name,
                `code`: Account code,
                `credit`: total amount of credit,
                `debit`: total amount of debit,
                `balance`: total amount of balance,
        """

        account_result = {}
        # Prepare sql query base on selected parameters from wizard
        tables, where_clause, where_params = self.env['account.move.line']._query_get()
        tables = tables.replace('"','')
        if not tables:
            tables = 'account_move_line'
        wheres = [""]
        if where_clause.strip():
            wheres.append(where_clause.strip())
        filters = " AND ".join(wheres)
        # compute the balance, debit and credit for the provided accounts
        request = ("SELECT account_id AS id, SUM(debit) AS debit, SUM(credit) AS credit, (SUM(debit) - SUM(credit)) AS balance" +\
                   " FROM " + tables + " WHERE account_id IN %s " + filters + " GROUP BY account_id")
        params = (tuple(accounts.ids),) + tuple(where_params)
        self.env.cr.execute(request, params)
        for row in self.env.cr.dictfetchall():
            account_result[row.pop('id')] = row

        account_res = []
        for account in accounts:
            res = dict((fn, 0.0) for fn in ['credit', 'debit', 'balance'])
            currency = account.currency_id and account.currency_id or account.company_id.currency_id
            res['code'] = account.code
            res['name'] = account.name
            if account.id in account_result:
                res['debit'] = account_result[account.id].get('debit')
                res['credit'] = account_result[account.id].get('credit')
                res['balance'] = account_result[account.id].get('balance')
            if display_account == 'all':
                account_res.append(res)
            if display_account == 'not_zero' and not currency.is_zero(res['balance']):
                account_res.append(res)
            if display_account == 'movement' and (not currency.is_zero(res['debit']) or not currency.is_zero(res['credit'])):
                account_res.append(res)
        return account_res

    @api.multi
    def print_excel_report(self):
        self.ensure_one()
        data = {}
        data['ids'] = self.env.context.get('active_ids', [])
        data['model'] = self.env.context.get('active_model', 'ir.ui.menu')
        data['form'] = self.read(['date_from', 'date_to', 'journal_ids', 'target_move', 'company_id','display_account'])[0]
        used_context = self._build_contexts(data)
        data['form']['used_context'] = dict(used_context, lang=self.env.context.get('lang') or 'en_US')
        accounts = self.env['account.account'].search([])
        read_lines = self._get_accounts(accounts, data['form'].get('display_account'))
        from pprint import pprint
        pprint(read_lines)
        fp = BytesIO()
        workbook = xlsxwriter.Workbook(fp)
        worksheet = workbook.add_worksheet('Aged Partner Balance Report')
        header_style = workbook.add_format({'font_name': 'Helvetica', 'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'border': 0})
        metric_style = workbook.add_format({'font': 'Helvetica', 'font_size': 10, 'bold': True,'align': 'center'})
        data_style = workbook.add_format({'font': 'Helvetica', 'font_size': 10,'align': 'center'})
        worksheet.merge_range('A2:I1', 'TRIAL BALANCE REPORT', header_style)
        row = 0
        col = 0
        row += 2
        worksheet.set_column('A:A', 10)
        worksheet.set_column('B:B', 30)
        worksheet.set_column('E:E', 20)
        worksheet.write(row,col, 'Display Account', metric_style)
        worksheet.write(row,1,self.display_account,data_style)
        worksheet.write(row,2,'Date From',metric_style)
        worksheet.write(row,3,str(self.date_from),data_style)
        worksheet.write(row,4,'Target Move',metric_style)
        worksheet.write(row,5,self.target_move,data_style)
        row +=1
        worksheet.write(row,4,'Date To',metric_style)
        worksheet.write(row,5,str(self.date_to),data_style)
        row +=2
#        worksheet.set_column('D:D', 20)
        worksheet.write(row,0,'Code',metric_style)
        worksheet.write(row,1,'Account',metric_style)
        worksheet.write(row,2,'Debit',metric_style)
        worksheet.write(row,3,'Credit',metric_style)
        worksheet.write(row,4,'Balance',metric_style)
        row +=1
        for line in read_lines:
            worksheet.write(row,0,line.get('code'))
            worksheet.write(row,1,line.get('name'), data_style)
            worksheet.write(row,2,line.get('debit'), data_style)
            worksheet.write(row,3,line.get('credit'), data_style)
            worksheet.write(row,4,line.get('balance'), data_style)
            row +=1
        workbook.close()
        report_id = self.env['account.report.excel'].create({
            'file_download': base64.encodestring(fp.getvalue()),
            'name': 'partner_ledger_xls''.xlsx'
        })
        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/?model=account.report.excel&id={}&field=file_download&filename_field=name&download=true'.format(report_id.id),
        }

