import bottlenose #https://github.com/lionheart/bottlenose library
from bs4 import BeautifulSoup
import ezodf
from lxml import html
import requests
import re
import linecache

#accessing spreadsheet:
spreadsheet = ezodf.opendoc("booklisy.ods") #open file
sheets = spreadsheet.sheets #init sheets
sheet= sheets['sure']
sheet['G1'].set_value("AWS_title")
sheet['H1'].set_value("AWS_authors")
sheet['I1'].set_value("AWS_publisher")
sheet['J1'].set_value("AWS_isbn")
sheet['K1'].set_value("AWS_url")

AWS_ACCESS_KEY_ID = ''	#your aws access key id
AWS_SECRET_ACCESS_KEY = '' #your aws secret access key
AWS_ASSOCIATE_TAG = '' #your aws associate tag
SERVICE_DOMAIN = 'IN'
amazon = bottlenose.Amazon(AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_ASSOCIATE_TAG,MaxQPS=0.9)
#other examples:
#response = amazon.ItemLookup(ItemId="0596520999", ResponseGroup="Images", SearchIndex="Books", IdType="ISBN")
#response = amazon.ItemSearch(Keywords="Kindle 3G", SearchIndex="All")


def backchodi (row,bname): 
	#for filtering; to remove and select search terms:
	bname=re.sub("[\,,\?,\:,\.,\;,\+,\']",' ',bname) #to remove , ? : . ; characters
	bname=re.sub(r'[^\x00-\x7F]+',' ',bname) #to remove non ASCII characters
	bname=re.sub("\d+", " ", bname)  #to remove numbers
	bname=re.sub('[!@#$]', ' ', bname)
	bname=re.sub(r'^"|"$', ' ', bname)
	bname=re.sub("([\(\[]).*?([\)\]])", "\g<1>\g<2>", bname) # to remove data within parentheses (,),[and]
	bname=re.sub('[\(,\),\[,\]]', " ", bname) #to remove parentheses (,)and[,]
	bname=re.sub('New', " ", bname) #to remove unwanted strings ('New') and replace it with space
	bname=re.sub('Delhi', " ", bname)
	bname=re.sub('new', " ", bname)
	bname=re.sub('delhi', " ", bname)
	bname=re.sub('York', " ", bname)
	bname=re.sub('york', " ", bname)
	bname=re.sub('Jersey', " ", bname)
	bname=re.sub('London', " ", bname)
	bname=re.sub('Tata', " ", bname)
	bname=re.sub('tata', " ", bname)
	bname=re.sub('elect', "electrical", bname)
	bname=re.sub('TMH', " ", bname)
	bname=re.sub('IEEE', " ", bname)
	bname=re.sub('Publication', " ", bname)
	bname=re.sub('MacMillan', " ", bname)
	bname=re.sub('Macmillan', " ", bname)
	bname=re.sub('macmillan', " ", bname)
	bname=re.sub('Mc', " ", bname)
	bname=re.sub('mc', " ", bname)
	bname=re.sub('graw', " ", bname)
	bname=re.sub('Publishing', " ", bname)
	bname=re.sub('Publishers', " ", bname)
	bname=re.sub('publishers', " ", bname)
	bname=re.sub('publishing', " ", bname)
	bname=re.sub('publications', " ", bname)
	bname=re.sub('Publications', " ", bname)
	bname=re.sub('Publication', " ", bname)
	bname=re.sub('publication', " ", bname)
	bname=re.sub('Graw', " ", bname)
	bname=re.sub('hill', " ", bname)
	bname=re.sub('Hill', " ", bname)
	bname=re.sub('John', " ", bname)
	bname=re.sub('john', " ", bname)
	bname=re.sub('Wiley', " ", bname)
	bname=re.sub('wiley', " ", bname)
	bname=re.sub('Sons', " ", bname)
	bname=re.sub('sons', " ", bname)
	bname=re.sub('Pvt', " ", bname)
	bname=re.sub('pvt', " ", bname)
	bname=re.sub('PVT', " ", bname)
	bname=re.sub('LTD', " ", bname)
	bname=re.sub('Ltd', " ", bname)
	bname=re.sub('ltd', " ", bname)
	bname=re.sub('Dr ', " ", bname)
	bname=re.sub('Learning', " ", bname)
	bname=re.sub('Cengage', " ", bname)
	bname=re.sub(' th Ed ', " ", bname)
	bname=re.sub(' nd Ed ', " ", bname)
	bname=re.sub(' rd Ed ', " ", bname)
	bname=re.sub(' st Ed ', " ", bname)
	bname=re.sub(' th ed ', " ", bname)
	bname=re.sub(' nd ed ', " ", bname)
	bname=re.sub(' rd ed ', " ", bname)
	bname=re.sub(' st ed ', " ", bname)
	bname=re.sub(' st ', " ", bname)
	bname=re.sub(' nd ', " ", bname)
	bname=re.sub(' rd ', " ", bname)
	bname=re.sub(' th ', " ", bname)
	bname=re.sub('Prentice', " ", bname)
	bname=re.sub('Vol', " ", bname)
	bname=re.sub('prentice', " ", bname)
	bname=re.sub('Pearson', " ", bname)
	bname=re.sub('pearson', " ", bname)
	bname=re.sub('Education', " ", bname)
	bname=re.sub('education', " ", bname)
	bname=re.sub('USA', " ", bname)
	bname=re.sub('UK', " ", bname)
	bname=re.sub('Age', " ", bname)
	bname=re.sub('Springer', " ", bname)
	bname=re.sub('Oxford', " ", bname)
	bname=re.sub('University', " ", bname)
	bname=re.sub('Hall', " ", bname)
	bname=re.sub('hall', " ", bname)
	bname=re.sub('India', " ", bname)
	bname=re.sub('PHI', " ", bname)
	bname=re.sub(' Co ', " ", bname)
	bname=re.sub('Michigan', " ", bname)
	bname=re.sub('South', " ", bname)
	bname=re.sub('West', " ", bname)
	bname=re.sub('North', " ", bname)
	bname=re.sub('East', " ", bname)
	bname=re.sub('Free', " ", bname)
	bname=re.sub('Press', " ", bname)
	bname=re.sub('edition', " ", bname)
	bname=re.sub('Edition', " ", bname)
	bname=re.sub('edtion', " ", bname)
	bname=re.sub('Dhanpat Rai', " ", bname)
	bname=re.sub('II', " ", bname)
	bname=re.sub('III', " ", bname)
	bname=re.sub('IV', " ", bname)
	bname=re.sub('Engg', "Engineering", bname) #to replace engg with engineering
	bname=re.sub('engg', "Engineering", bname)
	bname=re.sub('[&]', " ", bname)	#to replace & with space
	bname=re.sub('[\-]', " ", bname) #to replace - with space
	bname=re.sub('[\"]', " ", bname) #to replace " with space

	bname=re.sub(' +',' ',bname)	#to remove multiple spaces
	#bname= bname[:-15]	#to remove n characters from the end (to shorten search terms)
	bname= bname[:54]	#to select first n characters
	bname=re.sub(r'\b\w{1,3}\b','', bname)	#to remove single characters {1,n}, n length
	#print bname

	response = amazon.ItemSearch(Keywords=bname, SearchIndex="All", Region="IN")
	soup=BeautifulSoup(response)
	#print(soup.prettify())

	#for storing results in spreadsheet:
	
	for items in soup.find_all('item'):
		for isbn in items.asin: #store product(book) isbn in column 'J':
			print(isbn)
			sheet['J'+ str(row)].set_value(isbn)

		for url in items.detailpageurl: #store product url in column 'K':
			print(url)
			sheet['K'+ str(row)].set_value(url)

		allauthors=''
		for author in items.itemattributes.find_all('author'): #store Author in column 'H':
			if author.string:
				if (allauthors==''): 
					allauthors=(author.string)
				else: 
					allauthors=allauthors+';'+(author.string)
		print allauthors
		sheet['H'+ str(row)].set_value(allauthors)

		publisher=''
		for publishers in items.itemattributes.find_all('manufacturer'): #store manufacturer(Publisher) in column 'I':
			if publishers.string:
				if (publisher==''):
					publisher=(publishers.string)
				else: 
					publisher=publisher+';'+(publishers.string)
		print publisher
		sheet['I'+ str(row)].set_value(publisher)	

		title=''
		for titles in items.itemattributes.find_all('title'): #store book title in column 'G':
			if titles.string:
				if (title==''):
					title=(titles.string)
				else: 
					title=publisher+';'+(titles.string)
		print title
		sheet['G'+ str(row)].set_value(title)	

		break	#break is used here to shortlist first search result only!
	print ("\n")
	spreadsheet.save()	




for i in range(1,sheet.nrows()):	#open spreadsheet and pass search terms from column 'B' to function
	
	cell=sheet['B'+ str(i)]	

	if cell:
		celldata = str(cell.value.encode("utf-8"))

		if (sheet['G' + str(i)]):
			print ''
		else:
			backchodi(i,celldata)
			print i,'of',sheet.nrows(),'done'

spreadsheet.save()