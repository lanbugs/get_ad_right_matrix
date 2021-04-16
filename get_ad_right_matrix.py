#!/usr/bin/env python3

#
# get_ad_right_matrix.py
# Export AD User -> Group Matrix to Excel
# Written by Maximilian Thoma 2021
#

import json
import re
import ldap3
import pandas as pd

########################################################################################################################
# NOTE:
# -----
# Following packages must be installed in your python environment:
# pandas, xslxwriter, ldap3
#
# Just install them with:
# pip install pandas xslxwriter, ldap3
#
########################################################################################################################
# Settings

# LDAP server ip or fqdn
LDAP_SERVER = '10.1.1.231'
# LDAP port 389 = unencrypted, 636 = encrypted
PORT = 389
# Use SSL? True/False
USE_SSL = False
# LDAP bind user DN
BIND = 'CN=ldap bind,CN=Users,DC=lab,DC=local'
# LDAP bind user password
BIND_PW = 'Test12345!'
# Base search DN
SEARCH = 'OU=lab,DC=lab,DC=local'
# All users regardless deactivated or activated
SEARCH_FILTER = '(&(objectclass=user)(sAMAccountName=*))'
# All users who are not deactivated
#SEARCH_FILTER = '(&(objectclass=user)(sAMAccountName=*)(!(UserAccountControl:1.2.840.113556.1.4.803:=2)))'
# All users who are not deactivated and in special group
#SEARCH_FILTER = '(&(objectclass=user)(sAMAccountName=*)(!(UserAccountControl:1.2.840.113556.1.4.803:=2))(memberOf=CN=b_testgruppe und restlicher DN))'
# Output file
FILE = 'output.xlsx'
########################################################################################################################

def main():
    # Connect to LDAP and query
    server = ldap3.Server(LDAP_SERVER, port=389, use_ssl=USE_SSL)
    conn = ldap3.Connection(server, BIND, BIND_PW, auto_bind=True)
    conn.search(SEARCH, SEARCH_FILTER, attributes=['memberOf', 'sAMAccountName'])
    response = json.loads(conn.response_to_json())

    def get_cn(cn_str):
        cn = re.findall(r"CN=([^,]*),?", cn_str)[0]
        return cn

    buffer_users = {}
    buffer_user_in_group = {}

    for entry in response['entries']:
        # Get short and long username
        long_username = get_cn(entry['dn'])
        short_username = entry['attributes']['sAMAccountName'].lower()

        # append to users dir
        buffer_users[short_username] = long_username

        # go trough groups
        for group in entry['attributes']['memberOf']:
            # add to group buffer
            group_name = get_cn(group)
            if group_name not in buffer_user_in_group:
                buffer_user_in_group[group_name] = []
            if short_username not in buffer_user_in_group[group_name]:
                buffer_user_in_group[group_name].append(short_username)

    matrix = {}
    length_cell = 0

    for group, users in buffer_user_in_group.items():
        matrix[group] = {}

        for user, long_user in buffer_users.items():
            index = "%s - %s" % (user, long_user)
            # determine width of 1 column
            index_length = len(index)
            if index_length > length_cell:
                length_cell = index_length

            if user in users:
                matrix[group][index] = "X"
            else:
                matrix[group][index] = "-"

    # generate data matrix with pandas
    a = pd.DataFrame(matrix)

    # create excel file
    writer = pd.ExcelWriter(FILE, engine='xlsxwriter')

    # write pandas matrix to sheet1
    a.to_excel(writer, sheet_name="Sheet1", startrow=1, header=False)

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # format header line
    header_format = workbook.add_format(
        {
            'bold': True,
            'valign': 'bottom',
            'fg_color': '#D7E4BC',
            'border': 1,

        }
    )
    # set header line text rotation to 90 degree
    header_format.set_rotation(90)

    # apply header format
    for col_num, value in enumerate(a.columns.values):
        worksheet.write(0, col_num + 1, value, header_format)

    # format for X cells
    format2 = workbook.add_format(
        {
            'bg_color': '#C6EFCE',
            'font_color': '#006100'
        }
    )

    # set autofilter in first line
    cols_count = len(a.columns.values)
    worksheet.autofilter(0, 0, 0, cols_count)

    # set column width
    worksheet.set_column(0, 0, length_cell+1)
    worksheet.set_column(1, cols_count, 3)

    # freeze panes
    worksheet.freeze_panes(1, 1)

    # conditional formatting
    worksheet.conditional_format('A1:ZA65535', {
        'type': 'cell',
        'criteria': '=',
        'value': '"X"',
        'format': format2
    })

    # save excel file
    writer.save()


if __name__ == "__main__":
    main()
