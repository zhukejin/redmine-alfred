# coding=UTF-8
import sys
import MyRedmine

rest_key = 'you key' # 从 redmine 后台可以看到

#  创建实例
my_redmine = MyRedmine.RedmineQuery(rest_key)

# query
query = sys.argv[1] if len(sys.argv) >= 2 else ''

# 输入为0 ，进行本月个人加班统计
if query == '0':
    my_redmine.get_person_hour()
else:
    my_redmine.get_issue_info(query)
