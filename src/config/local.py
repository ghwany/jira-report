from datetime import datetime, timedelta, timezone
import json, sys, os

now = datetime.now(tz=timezone.utc)

START_DATE = now - timedelta(weeks=1)
END_DATE = now + timedelta(days=1)

JIRA_ID = "*****"
JIRA_TOKEN = "*****"
JIRA_SERVER = "https://bighitcorp.atlassian.net"
JIRA_REQ_URL = '/browse/'
JIRA_PROJECT = "WEVSEC"
JIRA_PROJECT_STATUS = ['Open', 'TO DO', '작업 중', 'REPORTING', 'IN REVIEW', 'WAITING FOR RESPONSE', '완료', 'BLOCKED']
JIRA_JQL = 'status in ({}) AND ' \
           'type in (Task, Sub-task) ' \
           'ORDER BY assignee ASC, component ASC, status ASC, created DESC'.format(', '.join(f'"{s}"' for s in JIRA_PROJECT_STATUS))
# Key: 별칭, Value[Array]: 별칭으로 변경할 상태 값들
JIRA_ALIAS_ISSUE_STATUS = {
    '대기중': ['Waiting for Response', 'Blocked'],
    '할 일': ['해야 할 일', '열기'],
    '검토요청': ['IN REVIEW', 'Reporting'],
}

NO_COMPONENT = '미분류'


def get_jira_alias_issue_status(status: str):
    return next(
        (k for k, v in JIRA_ALIAS_ISSUE_STATUS.items() if status in v),
        status
    )


# Secrets load
with open(os.sep.join(__file__.split(os.sep)[:-1] + ['secrets.json'])) as fs:
    secrets = json.loads(fs.read())
    for key, value in secrets.items():
        setattr(sys.modules[__name__], key, value)
