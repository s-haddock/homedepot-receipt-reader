import imaplib, email, re, os, csv
from BeautifulSoup import BeautifulSoup
from datetime import datetime
import simplejson as json

print """*******************************************************
PROGRAM NAME: Home Depot Receipt Reader

PROGRAMMER: S. Haddock (sjay89@gmail.com)

CREATED DATE: 1/30/2018

USAGE: The email you enter must have enable less secure apps turned on. 
	   https://support.google.com/a/answer/6260879?hl=en

DESCRIPTION: This program logs into your email, searches and then downloads 
	all of the receipts found in the date range into a folder. It then
	reads through each receipt extracting data. There are three files that are 
	saved...
	
	line_items.csv contains all of the combined line items from the receipt
	receipt_data contains all of the data for each receipt
	receipt_data.json holds all of the data in json format
*******************************************************

"""
 # Start Date Input
while True:
	try:
		start_date = raw_input('Enter the start date in MM/DD/YYYY format: ')

		start_date = datetime.strptime(start_date, '%m/%d/%Y')
		break
	except:
		print 'Invalid entry format must match MM/DD/YYYY...'

# End Date Input
while True:
	try:
		end_date = raw_input('\nEnter the end date in MM/DD/YYYY format: ')

		end_date = datetime.strptime(end_date, '%m/%d/%Y')
		break
	except:
		print 'Invalid entry format must match MM/DD/YYYY...\n'

#Email Authentication
while True:
	try:
		# Username Input
		username = raw_input('\nEnter your email address: ')

		# Password Input
		password = raw_input('\nEnter your email password: ')

		client = imaplib.IMAP4_SSL('imap.gmail.com')
		client.login(username, password)
		break
	except Exception as e:
		print e.message

# Load Mailboxes
status, mailboxes = client.list()
if not status == 'OK':
    raise Exception('Failed to get mailboxes.')

# Prompt for inbox name
while True:
	try:
		# Username Input
		inbox = raw_input('\nEnter the email folder name you wish to search: ')

		status, data = client.select(inbox)
		if not status == 'OK':
		    raise Exception('Failed to select INBOX.')
		break
	except Exception as e:
		print e.message

inbox = 'INBOX'
status, data = client.select(inbox)
if not status == 'OK':
    raise Exception('Failed to select INBOX.')

# Searching for emails
print '\nSearching ' + inbox + ' for emails from HomeDepotReceipt@homedepot.com'
status, email_ids = client.uid('search', None, '(FROM "HomeDepotReceipt@homedepot.com")')
if not status == 'OK':
    raise Exception('There was an error searching ' + inbox + ' for Home Depot Receipts')

# Check number of emails
email_ids = email_ids[0].split()
if len(email_ids) == 0:
    raise Exception('No emails found in the date range provided...')

print 'Found a total of ', len(email_ids), ' emails in the folder ', inbox



# Create Email Folder and clean dir
cwd = os.getcwd()

if not os.path.isdir(cwd + '/Home Depot Receipts'):
	os.mkdir('Home Depot Receipts')

homeDepotRoot = cwd + '/Home Depot Receipts/'

dirs = os.listdir(homeDepotRoot)
for file in dirs:
	if not os.path.isdir(homeDepotRoot + file):
		os.remove(homeDepotRoot + file)

if not os.path.isdir(homeDepotRoot + 'Receipts'):
	os.mkdir(homeDepotRoot + 'Receipts')

receiptFolder = homeDepotRoot + 'Receipts/'

dirs = os.listdir(receiptFolder)
for file in dirs:
	if not os.path.isdir(receiptFolder + file):
		os.remove(receiptFolder + file)

output = []
errors = []

for email_id in email_ids:
	try:
		status, data = client.uid('fetch', email_id, "(RFC822)")
		if not status == 'OK':
			raise Exception(status)

		raw_email = data[0][1]

		msg = email.message_from_string(raw_email)

		date = msg['Date'].split('+')[0].strip() if '+' in msg['Date'] else msg['Date'].split('-')[0].strip()
		parsedDate = datetime.strptime(date, "%a, %d %b %Y %H:%M:%S")
		date = parsedDate.strftime('%m/%d/%Y')
		parsedDate = datetime.strptime(date, "%m/%d/%Y")

		if parsedDate < start_date:
			print 'Email Id -', email_id, 'was received on', date, 'it is outside of the provided date range. Skipping...'
			continue

		if parsedDate > end_date:
			print 'Email Id -', email_id, 'was received on', date, 'it is outside of the provided date range. Ending search...'
			break

		for part in msg.walk():
			if part.get_content_type() == 'text/html':
				html = part.get_payload(decode=True)
				soup = BeautifulSoup(html)

				# find <pre> tag which wraps receipt body
				pre = soup.find('pre')

				# find div rows
				divs = [div for div in pre.findAll('div') if not div.text == '&nbsp;' and not div.text == None]
				
				email_data = {'line_items':[], 'email_id': email_id}
				more_line_items = True
				for n in range(len(divs)):
					div = divs[n]
					if len(div.text.split('SUBTOTAL')) > 1:
						more_line_items = False

					transaction_date = re.findall('\d\d[/]\d\d[/]\d\d\s+\d\d[:]\d\d\s+\D\D', div.text)
					if not email_data.has_key('transaction_date') and len(transaction_date) > 0:
						email_data['transaction_date'] = transaction_date[0] 

					item_description = div.text.split('&lt;A&gt;')
					if len(item_description) > 1 and more_line_items:
						line_items = []
						for i in range(0, 5):
							line_items.append(divs[n + i].text)

							next_line = divs[n + i + 1].text
							
							if len(next_line.split('&lt;A&gt;')) > 1 or len(next_line.split('SUBTOTAL')) > 1:
								break

						line_data = {'total': None, 'item_code': None, 'description': None, 'quantity': None, 'price_per_unit': None}

						for l in range(len(line_items)):
							line = line_items[l]

							# Line Total
							total = re.findall('\d+[.]\d\d$', line)
							if len(total) > 0:
								line_data['total'] = float(total[0].strip())

							#Quantity and Price Per Item
							quantityAndPrice = re.findall('\d+[@]\d+[.]\d\d', line)
							if len(quantityAndPrice) > 0:
								line_data['quantity'] = int(quantityAndPrice[0].split('@')[0])
								line_data['price_per_unit'] = float(quantityAndPrice[0].split('@')[1])

							# Item Code
							asymbol = line.split('&lt;A&gt;')
							if len(asymbol) > 1:
								line_data['item_code'] = asymbol[0].strip()
								if len(line_items) > l+1 and not len(re.findall('\d+[@]\d+[.]\d\d', line_items[l+1])) > 0:
									line_data['description'] = line_items[l+1]

						if line_data['quantity'] == None:
							line_data['quantity'] = 1
							line_data['price_per_unit'] = line_data['total']

						email_data['line_items'].append(line_data)

					# Subtotal
					subtotal = re.findall('SUBTOTAL', div.text)
					if not email_data.has_key('subtotal') and len(subtotal) > 0:
						subtotal = re.findall('\d+[.]\d\d$', div.text)
						if len(subtotal) > 0:
							email_data['subtotal'] = float(subtotal[0].strip())

					# Sales tax
					sales_tax = re.findall('SALES TAX', div.text)
					if not email_data.has_key('sales_tax') and len(sales_tax) > 0:
						sales_tax = re.findall('\d+[.]\d\d$', div.text)
						if len(sales_tax) > 0:
							email_data['sales_tax'] = float(sales_tax[0].strip())

					# Total Price
					total_price = re.findall('TOTAL', div.text)
					if not email_data.has_key('total_price') and len(total_price) > 0 and len(subtotal) == 0:
						total_price = re.findall('\d+[.]\d\d$', div.text)
						if len(total_price) > 0:
							email_data['total_price'] = float(total_price[0].strip())

				output.append(email_data)
				f = open('%s/%s.mht' %(receiptFolder, email_id), 'wb')
				print "Saving email id -", email_id, ' Transaction Date: ', email_data['transaction_date']
				f.write(raw_email)
				break

	except Exception as e:
		print e.message
		errors.append(email_id)

line_items = [['Email Id', 'Receipt Date', 'Item Code', 'Description', 'Quantity', 'Price Per Unit', 'Total']]

for receipt in output:
	for line_item in receipt['line_items']:
		line_items.append([
			receipt['email_id'],
			receipt['transaction_date'], 
			line_item['item_code'], 
			line_item['description'],
			line_item['quantity'],
			line_item['price_per_unit'],
			line_item['total']
		])

with open(homeDepotRoot + 'line_items.csv', 'wb') as writeFile:
    writer = csv.writer(writeFile, delimiter=',')
    writer.writerows(line_items)

receipt_data = [['Email Id', 'Receipt Date', 'Subtotal', 'Sales Tax', 'Total']]

sum_subtotal = 0
sum_tax = 0
sum_total = 0

for receipt in output:
	receipt_data.append([
		receipt['email_id'],
		receipt['transaction_date'], 
		receipt['subtotal'] if receipt.has_key('subtotal') else None, 
		receipt['sales_tax'] if receipt.has_key('sales_tax') else None,
		receipt['total_price'] if receipt.has_key('total_price') else None
	])

	sum_subtotal += float(receipt['subtotal'] if receipt.has_key('subtotal') else 0)
	sum_tax += float(receipt['sales_tax'] if receipt.has_key('sales_tax') else 0)
	sum_total += float(receipt['total_price'] if receipt.has_key('total_price') else 0)

receipt_data.append([])
receipt_data.append([])
receipt_data.append(['SUMMARY TOTALS','', sum_subtotal, sum_tax, sum_total])


with open(homeDepotRoot + 'receipt_data.csv', 'wb') as writeFile:
    writer = csv.writer(writeFile, delimiter=',')
    writer.writerows(receipt_data)

	
## JSON file includes all data
outfile = open(homeDepotRoot + 'receipt_data.json', "w")
outfile.write(json.dumps(output, indent=4, sort_keys=True))
outfile.close()

if len(errors) > 0:
	print 'The following email ids encountered errors and were not added to the output files...'
	print errors

raw_input('\nPress enter to quit...')