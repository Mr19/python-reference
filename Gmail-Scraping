#####
# Written by Gaddy
#####

#!/usr/bin/env python

import sys
import imaplib
import email
import email.header
import getpass
import re
import gdata.spreadsheet.service
import datetime

EMAIL_ACCOUNT = '<EmailID>'
password = getpass.getpass()
mail = imaplib.IMAP4_SSL('imap.gmail.com')
	
def process_mailbox(mail):
		mail.select("<EmailFolder>") # Folder to scan. All emails under this folder will be scanned and scraped
		result, data = mail.search(None, "ALL")
		if result != 'OK':
			print "No Messages Found!"
			return
		ids = data[0]
		id_list = ids.split()
		latest_email_id = id_list[-1]
		for latest_email_id in data[0].split():
			result, data = mail.fetch(latest_email_id, "(RFC822)")
			if result != 'OK':
				print "Error Getting Message", latest_email_id
				return
			msg = email.message_from_string(data[0][1])
			decode = email.header.decode_header(msg['Subject'])[0]
			subject = unicode(decode[0])
			subject = subject.replace("Fwd: ", "")
			subject = subject.replace("Re: ", "")
			#print subject
			#print 'Raw Date:', msg['Date']
			# Now convert to local date-time
			date_tuple = email.utils.parsedate_tz(msg['Date'])
			if date_tuple:
				local_date = datetime.datetime.fromtimestamp(
                email.utils.mktime_tz(date_tuple))
				#print "Local Date:", \
                #local_date.strftime("%a, %d %b %Y %H:%M:%S")
			raw_email = data[0][1]
			raw_email = raw_email.replace("\n", "")
			raw_email = raw_email.replace("\r", "")
			a = re.search("FIRST_NAME: ", raw_email)
			o = re.search("--=20", raw_email)
			need_email = raw_email[a.end():o.start()]
			#print need_email
			# replace below with anything you want to scrap in the email
			b = re.search("LAST_NAME: ", need_email)
			c = re.search("CITY: ", need_email)
			d = re.search("STATE: ", need_email)
			e = re.search("PHONE: ", need_email)
			f = re.search("PHONE_NUMBER: ", need_email)
			g = re.search("EMAIL_ADDRESS: ", need_email)
			h = re.search("COMMENTS: ", need_email)
			i = re.search("INVESTIGATION: ", need_email)
			j = re.search("INVESTIGATION2: ", need_email)
			k = re.search("AGREEMENT: ", need_email)
			l = re.search("IP: ", need_email)
			m = re.search("PAGE:", need_email)
			n = re.search("REFERER:", need_email)
			firstname = need_email[:b.start()]
			firstname = firstname.replace(">", "")
			#print firstname
			if c:
				lastname = need_email[b.end():c.start()]
				lastname = lastname.replace(">", "")
				city = need_email[c.end():d.start()]
				city = city.replace(">", "")
				#print lastname
				#print city
			else:
				lastname = need_email[b.end():d.start()]
				lastname = lastname.replace(">", "")
				city = "NA"
				#print lastname
				#print city
			if e:
				state = need_email[d.end():e.start()]
				state = state.replace(">", "")
				phone = need_email[e.end():g.start()]
				phone = phone.replace("-", "")
				phone = phone.replace(" ", "")
				phone = phone.replace(">", "")
				#print state
				#print phone
			else:
				state = need_email[d.end():f.start()]
				state = state.replace(">", "")
				phone = need_email[f.end():g.start()]
				phone = phone.replace("-", "")
				phone = phone.replace(" ", "")
				phone = phone.replace(">", "")
				#print state
				#print phone
			if h:
				emailid = need_email[g.end():h.start()]
				emailid = emailid.replace(">", "")
				comment = need_email[h.end():k.start()]
				comment = comment.replace(">", "")
				#print emailid
				#print comment
			else:
				comment = "NA"
			if i:
				emailid = need_email[g.end():i.start()]
				emailid = emailid.replace(">", "")
				comment = "NA"
				inv = need_email[i.end():k.start()]
				inv = inv.replace(">", "")
				#print emailid
				#print comment
				#print inv
			else:
				inv = "NA"
			if j:
				emailid = need_email[g.end():j.start()]
				emailid = emailid.replace(">", "")
				comment = "NA"
				inv = "NA"
				inv2 = need_email[j.end():k.start()]
				inv2 = inv2.replace(">", "")
				#print emailid
				#print comment
				#print inv
				#print inv2
			else:
				inv2 = "NA"
				#print inv2
			if k:
				agree = need_email[k.end():l.start()]
				agree = agree.replace(">", "")
				#print agree
			else:
				agree = "NA"
				#print agree
			if l:
				ip = need_email[l.end():m.start()]
				ip = ip.replace(">", "")
				#print ip
			else:
				ip = "NA"
				#print ip
			if m:
				page = need_email[m.end():n.start()]
				page = page.replace(">", "")
				page = page.replace(" ", "")
				#print page
			else:
				page = "NA"
				#print page
			if l:
				referer = need_email[n.end():]
				referer = referer.replace(">", "")
				referer = referer.replace(" ", "")				
				#print referer
			else:
				referer = "NA"
				#print referer
			spreadsheet_key = '<GoogleSheetKey>' # Add the key to which the data should be uploaded
			worksheet_id = '<WorksheetID>'	# Add worksheet id here		
			sheet = gdata.spreadsheet.service.SpreadsheetsService()
			#sheet.debug = False
			sheet.email = EMAIL_ACCOUNT
			sheet.password = password
			sheet.source = 'Python Update'
			sheet.ProgrammaticLogin()
			data = {}
			data['date'] = local_date.strftime("%a, %d %b %Y %H:%M:%S")
			data['subject'] = subject
			data['firstname'] = firstname
			data['lastname'] = lastname
			data['city'] = city
			data['state'] = state
			data['phone'] = phone
			data['email'] = emailid
			data['comments'] = comment
			data['investigation'] = inv
			data['investigation2'] = inv2
			data['agreement'] = agree
			data['ip'] = ip
			data['page'] = page
			data['referer'] = referer
			#print data

			entry = sheet.InsertRow(data, spreadsheet_key, worksheet_id)
			#if isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
			#	print "Insert row successful."
			#else:
			#	print "Insert row failed."
		print "\n %s Email(s) Updated" % (latest_email_id)
		
try:
    result, data = mail.login(EMAIL_ACCOUNT, password)
	
except imaplib.IMAP4.error:
    print "\nForgot your Password???...LOGIN FAILED!!! "
    sys.exit(1)		

if result == 'OK':
    process_mailbox(mail)
    mail.close()

mail.logout()
