# -*- coding: utf-8 -*-

# TODO 100人程度送った時点でConnectionが切断されるため、再接続処理の実装が求められる

# Author Tacke (tacke.jp@gmail.com)

###### コードブラッシュアップの指針 ######
# 大きな部分に分類 (設定ファイル読み込み, CSV読み込み, メール形式に変換, Gmailとの接続確立)
# 各行程で異常系を正しく作る
# テストを書いて品質を保証

import smtplib
import csv
import sys
from email.MIMEText import MIMEText
from email.Header import Header
from email.Utils import formatdate
from getpass import getpass

if __name__ == '__main__':

    argv = sys.argv
    if (len(argv) != 4):
        print 'usage: python %s <CSV> <SUBJECT> <CONTENTFILE>' % argv[0]
        quit()

    csvfile = argv[1]
    subject_u = argv[2].decode('utf8') # depend on terminal
    print subject_u
    contentfile = argv[3]
    content_u = file(contentfile).read().decode('utf8')

    #********** Settings **********

    # utf8-encoded tab-delimited csv 
    # the first column is address and others are inserted strings

    encoding = 'ISO-2022-JP'
    from_addr = 'meiwa.h20@gmail.com'
    fromtxt_u = u'明和学年会H20卒'
    #subject_u = u'明和学年同窓会のご案内'

    smtp_user = from_addr
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587

    #******************************

    reader = csv.reader(file(csvfile, 'r'), delimiter="\t")
    password = getpass()

    s = smtplib.SMTP(smtp_server, smtp_port)
    s.ehlo()
    s.starttls()
    s.ehlo()
    s.login(smtp_user, password)

    for row in reader:
        to_addr = row[0]
        insert_strs = map((lambda s: unicode(s, 'utf-8', 'ignore')), row[1:])

        msg = MIMEText(content_u % tuple(insert_strs), 'plain', encoding)
        msg['Subject'] = Header(subject_u, encoding)
        msg['From'] = Header(fromtxt_u, encoding)
        msg['To'] = to_addr
        msg['Date'] = formatdate()

        s.sendmail(from_addr, [to_addr], msg.as_string())

        print 'sent to ' + to_addr

    s.close()
    print 'finished!'
