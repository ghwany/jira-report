# -*- coding: utf-8 -*-

from jira import JIRA, JIRAError
from datetime import datetime
import pandas as pd
import xlsxwriter
import re
import argparse

from config.local import *


def get_comments(issue_comments, from_date, to_date):
    result_comment = []

    if issue_comments:

        for comment in issue_comments:
            comment_updated_date = datetime.strptime(comment.updated, '%Y-%m-%dT%H:%M:%S.%f%z')
            comment_created_date = datetime.strptime(comment.created, '%Y-%m-%dT%H:%M:%S.%f%z')
            comment_time = ""

            if from_date <= comment_updated_date <= to_date:
                comment_time = comment.updated
            elif from_date <= comment_created_date <= to_date:
                comment_time = comment.created

            if comment_time:
                comment_data = {
                    'comment_id': comment.id,
                    'comment_updated': datetime.strptime(comment_time, '%Y-%m-%dT%H:%M:%S.%f%z').strftime('%Y-%m-%d'),
                    'comment_updateAuthor': str(comment.updateAuthor),
                    'comment_body': str(comment.body)
                }
                result_comment.append(comment_data)

    return result_comment


def get_worklogs(issue_worklogs, from_date, to_date):
    result_comment = []

    if issue_worklogs:

        for worklog in issue_worklogs:
            worklog_updated_date = datetime.strptime(worklog.updated, '%Y-%m-%dT%H:%M:%S.%f%z')
            worklog_created_date = datetime.strptime(worklog.created, '%Y-%m-%dT%H:%M:%S.%f%z')
            worklog_time = ""

            if from_date <= worklog_updated_date <= to_date:
                worklog_time = worklog.updated
            elif from_date <= worklog_created_date <= to_date:
                worklog_time = worklog.created

            if worklog_time:
                worklog_data = {
                    'worklog_id': worklog.id,
                    'worklog_updated': datetime.strptime(worklog_time, '%Y-%m-%dT%H:%M:%S.%f%z').strftime('%Y-%m-%d'),
                    'worklog_updateAuthor': str(worklog.updateAuthor),
                    'worklog_timeSpent': worklog.timeSpent,
                    'worklog_body': str(worklog.comment)
                }
                result_comment.append(worklog_data)

    return result_comment


def set_worksheet_header(worksheet, is_write_reporter: bool = True, components: set = None):
    # Title Header
    hcol, hrow = 0, 0
    worksheet.write(
        hrow, hcol, 'Jira Summary {} - {}'.format(
            START_DATE.strftime('%Y-%m-%d'),
            END_DATE.strftime('%Y-%m-%d')),
        header)
    hrow += 1

    corr_col = 0 if is_write_reporter else 2

    if components:
        worksheet.write(hrow, hcol, '??????', table_header)
        worksheet.merge_range(hrow, hcol + 1, hrow, hcol + 8 - corr_col, '?????? ?????? ??????', table_header)
        hrow += 1
        for c in components:
            worksheet.write(hrow, hcol, c)
            worksheet.merge_range(hrow, hcol + 1, hrow, hcol + 8 - corr_col, '')
            hrow += 1
        hrow += 1

    # Column Header
    worksheet.merge_range(hrow, hcol, hrow + 1, hcol, '??????', table_header)
    if is_write_reporter:
        worksheet.merge_range(hrow, hcol + 1, hrow + 1, hcol + 1, '?????????', table_header)
        worksheet.merge_range(hrow, hcol + 2, hrow + 1, hcol + 2, '?????? ?????? ??????', table_header)
    worksheet.merge_range(hrow, hcol + 3 - corr_col, hrow, hcol + 8 - corr_col, '?????? ?????? ??????', table_header)
    hrow += 1
    worksheet.write(hrow, hcol + 3 - corr_col, 'Jira ??????', table_header)
    worksheet.write(hrow, hcol + 4 - corr_col, '??????', table_header)
    worksheet.write(hrow, hcol + 5 - corr_col, '????????????', table_header)
    worksheet.write(hrow, hcol + 6 - corr_col, '??????????????????', table_header)
    worksheet.write(hrow, hcol + 7 - corr_col, '?????????', table_header)
    worksheet.write(hrow, hcol + 8 - corr_col, '????????????', table_header)
    hrow += 1


def set_worksheet_component(worksheet,
                            assignee: str, component: str,
                            row: int, col: int,
                            new_assignee_start_row: int,
                            is_write_reporter: bool = True):
    previous_row = row - 1

    if is_write_reporter:
        # Reporter
        total_cells = previous_row - new_assignee_start_row + 1
        if total_cells > 1:
            worksheet.merge_range(new_assignee_start_row, col + 1, previous_row, col + 1, assignee,
                                  cell_format=cell_format)
            worksheet.merge_range(new_assignee_start_row, col + 2, previous_row, col + 2, '',
                                  cell_format=cell_format)
        else:
            worksheet.write(new_assignee_start_row, col + 1, assignee, cell_format)
            worksheet.write(new_assignee_start_row, col + 2, '', cell_format)

    # Componenet
    component = component or NO_COMPONENT
    component_row_cnt = row - new_assignee_start_row

    if component_row_cnt > 1:
        worksheet.merge_range(new_assignee_start_row, col, previous_row, col, component,
                              cell_format=cell_format)
    else:
        worksheet.write(new_assignee_start_row, col, component, cell_format)


def set_worksheet_ticket(worksheet, ticket,
                         row: int, col: int,
                         cell_format,
                         is_write_reporter: bool = True):
    corr_col = 0 if is_write_reporter else 2
    comment_count = len(ticket['comments']) + len(ticket['worklogs'])

    if comment_count > 1:
        worksheet.merge_range(row, col + 3 - corr_col, row + comment_count - 1, col + 3 - corr_col, '')
        worksheet.write_url(row, col + 3 - corr_col, JIRA_SERVER + JIRA_REQ_URL + ticket['key'], string=ticket['key'],
                            cell_format=cell_format)
        worksheet.merge_range(row, col + 4 - corr_col, row + comment_count - 1, col + 4 - corr_col, ticket['status'], cell_format)
        worksheet.merge_range(row, col + 5 - corr_col, row + comment_count - 1, col + 5 - corr_col, ticket['summary'], cell_format)
    else:
        worksheet.write_url(row, col + 3 - corr_col, JIRA_SERVER + JIRA_REQ_URL + ticket['key'], string=ticket['key'],
                            cell_format=cell_format)
        worksheet.write(row, col + 4 - corr_col, ticket['status'], cell_format)
        worksheet.write(row, col + 5 - corr_col, ticket['summary'], cell_format)


def set_worksheet_ticket_comment(worksheet, comments,
                                 row: int, col: int,
                                 cell_format,
                                 corr_row: int = 0,
                                 is_worklog: bool = False,
                                 is_write_reporter: bool = True):
    comment_type = 'worklog' if is_worklog else 'comment'
    corr_col = 0 if is_write_reporter else 2
    index = corr_row
    for comment in comments:
        # Replace special characters in a comment_body
        comment[f'{comment_type}_body'] = re.sub(r'\*\*|\d\.|\*\#|\d\)|\-|\*', '', comment[f'{comment_type}_body'])

        worksheet.write(row + index, col + 6 - corr_col, comment[f'{comment_type}_updated'], cell_format)
        worksheet.write(row + index, col + 7 - corr_col, comment[f'{comment_type}_updateAuthor'].split('/')[0], cell_format)
        worksheet.write(row + index, col + 8 - corr_col, comment[f'{comment_type}_body'], cell_format)
        index += 1


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='JIRA ???????????? ?????? ??????')
    parser.add_argument('--reporter', '-r', required=False, help='?????? ????????????(????????? ??????)??? ?????? [?????????: ??????(??????)]')
    parser.add_argument('--separate', '-s', required=False, default=True, help='???????????? ???????????? ?????? [?????????: True]')
    parser.add_argument('--output', '-o', required=False, help='?????? ?????? ?????? ??????')
    parser.add_argument('--date-range', '-d', required=False, help='????????? ?????? ????????? ?????? ?????? [??????(YYYY-MM-DD),??????]')
    parser.add_argument('--jira-proj', '-p', required=False, default=JIRA_PROJECT, help='JIRA Project')
    parser.add_argument('--jira-url', '-u', required=False, default=JIRA_SERVER, help='JIRA Server URI')
    args = parser.parse_args()

    if args.jira_proj:
        JIRA_PROJECT = args.jira_proj
    JIRA_ID = JIRA_PROJECT_AUTH[JIRA_PROJECT]['ID']
    JIRA_TOKEN = JIRA_PROJECT_AUTH[JIRA_PROJECT]['TOKEN']
    JIRA_SERVER = args.jira_url
    JIRA_JQL = f'project = {JIRA_PROJECT} AND {JIRA_JQL}'

    if not args.date_range:
        START_DATE = now - timedelta(weeks=1)
        END_DATE = now + timedelta(days=1)
    else:
        start_date, end_date = args.date_range.split(',')
        START_DATE = datetime.strptime(start_date, '%Y-%m-%d').replace(tzinfo=timezone.utc)
        END_DATE = datetime.strptime(end_date, '%Y-%m-%d').replace(tzinfo=timezone.utc)

    EXPORT_REPORTER = args.reporter
    OUTPUT_FILE = args.output or f'../output/jira-{JIRA_PROJECT}-{EXPORT_REPORTER or "all"}-{int(now.timestamp())}.xlsx'
    SEPARATE_REPORTER_BY_SHEET = args.separate

    if EXPORT_REPORTER:
        EXPORT_REPORTER = EXPORT_REPORTER.split(',')

    try:
        jira = JIRA(server=JIRA_SERVER, basic_auth=(JIRA_ID, JIRA_TOKEN))
        jira_issues = jira.search_issues(JIRA_JQL, maxResults=False)

        result_issues = pd.DataFrame()
        result_comment = pd.DataFrame()

        for issue in jira_issues:
            issue_components = ""
            if issue.fields.components:
                issue_components = ", ".join(i.name for i in issue.fields.components)

            issue_updated_date = datetime.strptime(issue.fields.updated, '%Y-%m-%dT%H:%M:%S.%f%z')
            issue_created_date = datetime.strptime(issue.fields.created, '%Y-%m-%dT%H:%M:%S.%f%z')

            # ?????? ???????????? ????????? ???????????? ?????????, ?????? ?????? ????????? ??????, ??????????????? ????????? ???????????? ??????
            if not (START_DATE <= issue_updated_date <= END_DATE):
                continue
            assignee = str(issue.fields.assignee).split('/')[0]
            if EXPORT_REPORTER and not next((i for i in EXPORT_REPORTER if i in assignee), None):
                continue

            comments = get_comments(jira.comments(issue), START_DATE, END_DATE)
            worklogs = get_worklogs(jira.worklogs(issue=issue), START_DATE, END_DATE)

            data = {
                'id': issue.id,
                'key': issue.key,
                'self': issue.self,
                'components': str(issue_components),
                'summary': str(issue.fields.summary),
                'assignee': str(issue.fields.assignee),
                'status': str(issue.fields.status.name),
                'updated': str(issue_updated_date.strftime('%Y-%m-%d')),
                'created': str(issue_created_date.strftime('%Y-%m-%d')),
                'description': str(issue.fields.description),
                'comments': comments,
                'worklogs': worklogs,
            }

            result_issues = result_issues.append(data, ignore_index=True)

    except JIRAError as e:
        print(e.status, e.text)
        raise e

    if len(result_issues.key):
        workbook = xlsxwriter.Workbook(OUTPUT_FILE)
        worksheets = {}

        # Define the formats
        header = workbook.add_format(
            {'bold': True, 'bg_color': '#D8E4BC', 'align': 'center', 'valign': 'vcenter'})
        table_header = workbook.add_format({'bold': True, 'bg_color': '#fffae6', 'valign': 'vcenter', 'align': 'center'})
        cell_format = workbook.add_format({'valign': 'vcenter'})

        # Write the results
        component = None
        assignee = ''

        # header
        row, col = 0, 0

        new_assignee_start_row = 0
        data_start_row = 3

        if SEPARATE_REPORTER_BY_SHEET:
            for idx, ticket in result_issues.iterrows():
                # assignee
                ticket_assignee = ticket['assignee'].split('/')[0]
                component = ticket['components'] or NO_COMPONENT
                if ticket_assignee not in worksheets:
                    worksheet = workbook.add_worksheet(ticket_assignee)
                    worksheets[ticket_assignee] = {'worksheet': worksheet, 'components': set()}
                worksheets[ticket_assignee]['components'].add(component)

        for idx, ticket in result_issues.iterrows():
            # assignee
            ticket_assignee = ticket['assignee'].split('/')[0]
            # ?????? ???????????? ??????, next??? ???????????? ????????? ??????????????? ?????? ???????????? ???????????? ?????? ????????? ????????? ????????? ???????????? ??????
            if EXPORT_REPORTER and not next((i for i in EXPORT_REPORTER if i in ticket_assignee), None):
                continue

            # ????????? ?????? ????????? ???????????? ??????
            if SEPARATE_REPORTER_BY_SHEET:
                worksheet = worksheets[ticket_assignee]['worksheet']
                if 'row' in worksheets[ticket_assignee]:
                    row = worksheets[ticket_assignee]['row']
                else:
                    row = len(worksheets[ticket_assignee]['components']) + 5
                    set_worksheet_header(worksheet, not SEPARATE_REPORTER_BY_SHEET, worksheets[ticket_assignee]['components'])
            # ?????? ???????????? Summary ????????? ??????
            else:
                if 'Summary' not in worksheets:
                    row = 3
                    worksheet = workbook.add_worksheet('Summary')
                    worksheets['Summary'] = {'worksheet': worksheet}
                    set_worksheet_header(worksheet, not SEPARATE_REPORTER_BY_SHEET)
                else:
                    worksheet = worksheets['Summary']['worksheet']

            # Change status
            ticket['status'] = get_jira_alias_issue_status(ticket['status'])

            # Comment, Work log ?????? ?????? Ticket ??? ??????, 1??? ???????????? ???????????? ??????
            comment_count = len(ticket['comments']) + len(ticket['worklogs'])
            set_worksheet_ticket(worksheet, ticket, row, col, cell_format, not SEPARATE_REPORTER_BY_SHEET)

            # Comment, Work log ?????? ??????
            set_worksheet_ticket_comment(worksheet, ticket['comments'], row, col, cell_format,
                                         is_write_reporter=not SEPARATE_REPORTER_BY_SHEET)
            set_worksheet_ticket_comment(worksheet, ticket['worklogs'], row, col, cell_format,
                                         corr_row=len(ticket['comments']),
                                         is_worklog=True,
                                         is_write_reporter=not SEPARATE_REPORTER_BY_SHEET)

            is_draw_merge_column = component != ticket['components'] or assignee != ticket_assignee

            if 'row' not in worksheets[ticket_assignee]:
                # ????????? ??? ????????? ??? ??? ?????? ????????? ???????????? ?????? ?????? ?????? ?????? ????????? ????????? ?????? ?????? ????????? ???
                if SEPARATE_REPORTER_BY_SHEET and assignee != ticket_assignee and assignee in worksheets:
                    set_worksheet_component(
                        worksheets[assignee]['worksheet'],
                        assignee, component,
                        worksheets[assignee]['row'], col, new_assignee_start_row, not SEPARATE_REPORTER_BY_SHEET)
                new_assignee_start_row = row
            elif is_draw_merge_column:
                # ?????? ??????????????? ???????????? ???????????? ???????????? ????????? ????????? ?????? ?????? ????????? ???
                set_worksheet_component(
                    worksheet,
                    assignee, component,
                    row, col, new_assignee_start_row, not SEPARATE_REPORTER_BY_SHEET)
                new_assignee_start_row = row

            assignee = ticket_assignee
            component = ticket['components']
            row = row + (comment_count if comment_count > 1 else 1)

            # ????????? ???????????? ?????? ????????? ????????? ?????? worksheets??? ??????
            if SEPARATE_REPORTER_BY_SHEET:
                worksheets[ticket_assignee]['row'] = row

        # ????????? ????????? ????????? ???????????? ??????
        if assignee in worksheets:
            set_worksheet_component(
                worksheets[assignee]['worksheet'],
                assignee, component,
                row, col, new_assignee_start_row, not SEPARATE_REPORTER_BY_SHEET)

        workbook.close()
