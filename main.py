# -*- coding: utf-8 -*-

from jira import JIRA, JIRAError
import pandas as pd
import xlsxwriter
import re

from config import *


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


if __name__ == '__main__':
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

            # 이슈 업데이트 날짜가 포함되지 않거나, 날짜 안에 생성한 댓글, 작업로그가 없으면 포함하지 않음
            if not (START_DATE <= issue_updated_date <= END_DATE):
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

    if len(result_issues.key):
        workbook = xlsxwriter.Workbook(OUTPUT_FILE)

        # Define the formats
        row = 0
        col = 0
        header = workbook.add_format(
            {'bold': True, 'bg_color': '#D8E4BC', 'align': 'center', 'valign': 'vcenter'})
        table_header = workbook.add_format({'bold': True, 'valign': 'vcenter', 'align': 'center'})
        cell_format = workbook.add_format({'valign': 'vcenter'})

        # Write the results
        worksheet = workbook.add_worksheet('Summary')
        worksheet.write(
            row, col, 'Jira Summary {} - {}'.format(
                START_DATE.strftime('%Y-%m-%d'),
                END_DATE.strftime('%Y-%m-%d')),
            header)
        row += 1
        component = NO_COMPONENT
        assignee = ''
        index = 0

        # header
        worksheet.merge_range(row, col, row + 1, col, '구분', table_header)
        worksheet.merge_range(row, col + 1, row + 1, col + 1, '보고자', table_header)
        worksheet.merge_range(row, col + 2, row + 1, col + 2, '주간 업무 보고', table_header)
        worksheet.merge_range(row, col + 3, row, col + 8, '일일 보고 내용', table_header)
        row += 1
        worksheet.write(row, col + 3, 'Jira 티켓', table_header)
        worksheet.write(row, col + 4, '상태', table_header)
        worksheet.write(row, col + 5, '업무요약', table_header)
        worksheet.write(row, col + 6, '업데이트일자', table_header)
        worksheet.write(row, col + 7, '작성자', table_header)
        worksheet.write(row, col + 8, '보고내용', table_header)
        row += 1
        # 치환이 필요한 문자열
        # 서수(1., 2. etc), 특수문자(*#, **)

        compt_row_cnt = 0
        new_component_start_row = 0
        data_start_row = row
        new_assignee_start_row = 0

        for idx, ticket in result_issues.iterrows():
            comment_count = len(ticket['comments'])
            if comment_count <= 1:
                last_row = row
            else:
                last_row = row + comment_count - 1

            # Change status
            ticket['status'] = get_jira_alias_issue_status(ticket['status'])

            if comment_count > 1:
                worksheet.merge_range(row, col + 3, last_row, col + 3, '')
                worksheet.write_url(row, col + 3, JIRA_SERVER + JIRA_REQ_URL + ticket['key'], string=ticket['key'],
                                    cell_format=cell_format)
                worksheet.merge_range(row, col + 4, last_row, col + 4, ticket['status'], cell_format)
                worksheet.merge_range(row, col + 5, last_row, col + 5, ticket['summary'], cell_format)
            else:
                worksheet.write_url(row, col + 3, JIRA_SERVER + JIRA_REQ_URL + ticket['key'], string=ticket['key'],
                                    cell_format=cell_format)
                worksheet.write(row, col + 4, ticket['status'], cell_format)
                worksheet.write(row, col + 5, ticket['summary'], cell_format)

            index = 0
            for comment in ticket['comments']:
                # Replace special characters in a comment_body
                comment['comment_body'] = re.sub(r'\*\*|\d\.|\*\#|\d\)|\-|\*', '', comment['comment_body'])

                worksheet.write(row + index, col + 6, comment['comment_updated'], cell_format)
                worksheet.write(row + index, col + 7, comment['comment_updateAuthor'].split('/')[0], cell_format)
                worksheet.write(row + index, col + 8, comment['comment_body'], cell_format)
                index += 1
            for worklog in ticket['worklogs']:
                # Replace special characters in a comment_body
                worklog['worklog_body'] = re.sub(r'\*\*|\d\.|\*\#|\d\)|\-|\*', '', worklog['worklog_body'])

                worksheet.write(row + index, col + 6, worklog['worklog_updated'], cell_format)
                worksheet.write(row + index, col + 7, worklog['worklog_updateAuthor'].split('/')[0], cell_format)
                worksheet.write(row + index, col + 8, worklog['worklog_body'], cell_format)
                index += 1

            # assignee
            ticket_assignee = ticket['assignee'].split('/')[0]

            if row == data_start_row:
                assignee = ticket_assignee
                new_assignee_start_row = row
            elif assignee != ticket_assignee or component != ticket['components']:
                previous_row = row - 1
                total_cells = previous_row - new_assignee_start_row + 1

                if total_cells > 1:
                    worksheet.merge_range(new_assignee_start_row, col + 1, previous_row, col + 1, assignee,
                                          cell_format=cell_format)
                    worksheet.merge_range(new_assignee_start_row, col + 2, previous_row, col + 2, '',
                                          cell_format=cell_format)
                else:
                    worksheet.write(new_assignee_start_row, col + 1, assignee, cell_format)
                    worksheet.write(new_assignee_start_row, col + 2, '', cell_format)
                print('row: {} => s: {}, e: {}'.format(row, new_assignee_start_row, previous_row))

                new_assignee_start_row = row
                assignee = ticket_assignee

            if idx == result_issues.index[-1]:
                total_cells = last_row - new_assignee_start_row + 1

                if total_cells > 1:
                    worksheet.merge_range(new_assignee_start_row, col + 1, last_row, col + 1, ticket_assignee,
                                          cell_format=cell_format)
                    worksheet.merge_range(new_assignee_start_row, col + 2, last_row, col + 2, '',
                                          cell_format=cell_format)
                else:
                    worksheet.write(new_assignee_start_row, col + 1, ticket_assignee, cell_format)
                    worksheet.write(new_assignee_start_row, col + 2, '', cell_format)
                print('row: {} => s: {}, e: {} (last)'.format(row, new_assignee_start_row, row))

            # component
            if component == NO_COMPONENT:
                new_component_start_row = row
            elif component != ticket['components']:
                component_row_cnt = row - new_component_start_row

                if component_row_cnt > 1:
                    worksheet.merge_range(new_component_start_row, col, row - 1, col, component,
                                          cell_format=cell_format)
                else:
                    worksheet.write(new_component_start_row, col, component, cell_format)

                new_component_start_row = row

            if idx == result_issues.index[-1]:
                component_row_cnt = last_row - new_component_start_row + 1
                if ticket['components'] == '':
                    component = NO_COMPONENT
                else:
                    component = ticket['components']

                if component_row_cnt > 1:
                    worksheet.merge_range(new_component_start_row, col, last_row, col, component,
                                          cell_format=cell_format)
                else:
                    worksheet.write(new_component_start_row, col, component, cell_format)

            component = ticket['components']

            if comment_count > 1:
                row = row + comment_count
            else:
                row = row + 1

        workbook.close()
