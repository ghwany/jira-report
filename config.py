from datetime import datetime, timedelta, timezone

now = datetime.now(tz=timezone.utc)

START_DATE = now - timedelta(weeks=1)
END_DATE = now + timedelta(days=1)

JIRA_SERVER = "https://*****.atlassian.net"
JIRA_REQ_URL = '/browse/'
JIRA_ID = "*****"
JIRA_TOKEN = "*****"
JIRA_PROJECT = "*****"
JIRA_PROJECT_STATUS = ['Open', 'TO DO', '작업 중', 'REPORTING', 'IN REVIEW', 'WAITING FOR RESPONSE', '완료']
JIRA_JQL = 'project = {} AND ' \
           'status in ({}) AND ' \
           'type in (Task, Sub-task) ' \
           'ORDER BY component ASC, assignee ASC, status ASC, created DESC'.format(
    JIRA_PROJECT, ', '.join(f'"{s}"' for s in JIRA_PROJECT_STATUS))

# Key: 별칭, Value[Array]: 별칭으로 변경할 상태 값들
JIRA_ALIAS_ISSUE_STATUS = {
    '대기중': ['Waiting for Response', 'Blocked'],
    '할 일': ['해야 할 일', '열기'],
    '검토요청': ['IN REVIEW', 'Reporting'],
}

OUTPUT_FILE = 'weekly_report_{}.xlsx'.format(now.strftime('%Y-%m-%d'))
NO_COMPONENT = '미분류'


def get_jira_alias_issue_status(status: str):
    return next(
        (k for k, v in JIRA_ALIAS_ISSUE_STATUS.items() if status in v),
        status
    )