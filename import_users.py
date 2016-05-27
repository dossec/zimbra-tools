#!/usr/bin/env python
# -*- coding: utf-8 -*-


import paramiko
import xlrd

'''
zimbra 批量导入用户,excel包含四个字段:

地址  显示名称    姓   邮件限额

以上四个字段都是文本格式
'''

sucess = []
wrong = []

def remote_import(user, dis_name, first_name, quote):
    ssh_host = '192.168.1.67'
    ssh_user = 'user'
    ssh_pass = 'password'
    ssh_port = 22
    
    # 记录日志
    paramiko.util.log_to_file('paramiko.log')

    # 建立ssh连接
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(ssh_host, ssh_port, ssh_user, ssh_pass, timeout=3)

    # 执行远程命令

    try:
        stdin, stdout, stderr = ssh.exec_command(
            '/opt/zimbra/bin/zmprov ca {}@domain.com password displayName {} sn {} zimbraMailQuota {};echo $?' \
                .format(user, dis_name, first_name, quote))
        # print(stdout.read())
    except:
        print('命令执行错误！')
    else:
        result = stdout.readlines()
        error = stderr.readlines()
        # print(result)
        # print(error)
        # print('>>>>>>>>>>>')
        if result[-1] == '0\n':
            print('用户: {} 导入成功！'.format(user))
            sucess.append(user)
        elif result[-1] == '1\n':
            print('【结果】', result)
            print('【注意】 用户: {} 导入失败,数据格式有问题！'.format(user))
            wrong.append(user)

        else:
            if 'exists' in error[0]:
                print('【结果】', error)
                print('【注意】 用户: {} 导入失败,用户已存在！'.format(user))
                wrong.append(user)
            elif 'zimbraMailQuota must be a valid' in error[0]:
                print('【注意】 用户: {} 导入失败,容量格式问题！'.format(user))
            else:
                print('【注意】 用户: {} 导入失败,其它原因！'.format(user))
                wrong.append(user)
                # print(item.strip('\n'))
        ssh.close()


def import_user():
    data = xlrd.open_workbook('excel.xlsx', encoding_override='utf-8')
    table = data.sheets()[0]

    nrows = table.nrows
    # ncols = table.ncols

    col1 = table.col_values(0)
    col2 = table.col_values(1)
    col3 = table.col_values(2)
    col4 = table.col_values(3)
    for user, dis_name, first_name, quote in zip(col1, col2, col3, col4):
        # print(user,'\t' ,dis_name,'\t' ,first_name)
        data = {
            '帐号': user,
            '显示名称': dis_name,
            '姓名': first_name,
            '限额': str(quote)
        }
        print('现在导入数据 >>>', data)
        remote_import(user, dis_name, first_name, quote)
        print('\n')
    print('###### 本次计划导入的数据：{} 条 ######'.format(nrows))
    print('###### 其中导入成功的数据：{} 条 ######'.format(len(sucess)))
    print('###### 其中导入失败的数据：{} 条 ######'.format(len(wrong)))
    print('>>>导入失败的用户如下，请检查<<<：')
    for item in wrong:
        print(item)


if __name__ == '__main__':
    try:
        import_user()
    except Exception as e:
        print(e)
    else:
        pass
