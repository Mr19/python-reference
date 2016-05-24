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

EMAIL_ACCOUNT = 'email@gmail.com'
password = getpass.getpass()
mail = imaplib.IMAP4_SSL('imap.gmail.com')
	
def process_mailbox(mail):
		mail.select("EmailFolder/EmailSubFolder")
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
			date_tuple = email.utils.parsedate_tz(msg['Date'])
			if date_tuple:
				local_date = datetime.datetime.fromtimestamp(
                email.utils.mktime_tz(date_tuple))
			raw_email = data[0][1]
			raw_email = raw_email.replace("\n", "")
			raw_email = raw_email.replace("\r", "")
			a = re.search("FIRST_NAME: ", raw_email)
			o = re.search("--=20", raw_email)
			need_email = raw_email[a.end():o.start()]
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
			if c:
				lastname = need_email[b.end():c.start()]
				lastname = lastname.replace(">", "")
				city = need_email[c.end():d.start()]
				city = city.replace(">", "")
			else:
				lastname = need_email[b.end():d.start()]
				lastname = lastname.replace(">", "")
				city = "NA"
			if e:
				state = need_email[d.end():e.start()]
				state = state.replace(">", "")
				phone = need_email[e.end():g.start()]
				phone = phone.replace("-", "")
				phone = phone.replace(" ", "")
				phone = phone.replace(">", "")
			else:
				state = need_email[d.end():f.start()]
				state = state.replace(">", "")
				phone = need_email[f.end():g.start()]
				phone = phone.replace("-", "")
				phone = phone.replace(" ", "")
				phone = phone.replace(">", "")
			if h:
				emailid = need_email[g.end():h.start()]
				emailid = emailid.replace(">", "")
				comment = need_email[h.end():k.start()]
				comment = comment.replace(">", "")
			else:
				comment = "NA"
			if i:
				emailid = need_email[g.end():i.start()]
				emailid = emailid.replace(">", "")
				comment = "NA"
				inv = need_email[i.end():k.start()]
				inv = inv.replace(">", "")
			else:
				inv = "NA"
			if j:
				emailid = need_email[g.end():j.start()]
				emailid = emailid.replace(">", "")
				comment = "NA"
				inv = "NA"
				inv2 = need_email[j.end():k.start()]
				inv2 = inv2.replace(">", "")
			else:
				inv2 = "NA"
			if k:
				agree = need_email[k.end():l.start()]
				agree = agree.replace(">", "")
			else:
				agree = "NA"
			if l:
				ip = need_email[l.end():m.start()]
				ip = ip.replace(">", "")
			else:
				ip = "NA"
			if m:
				page = need_email[m.end():n.start()]
				page = page.replace(">", "")
				page = page.replace(" ", "")
			else:
				page = "NA"
			if l:
				referer = need_email[n.end():]
				referer = referer.replace(">", "")
				referer = referer.replace(" ", "")				
			else:
				referer = "NA"
			spreadsheet_key = 'google_sheet_key'
			worksheet_id = 'od6'			
			sheet = gdata.spreadsheet.service.SpreadsheetsService()
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
			entry = sheet.InsertRow(data, spreadsheet_key, worksheet_id)
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
