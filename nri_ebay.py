import bs4, xlsxwriter, requests, random
import time
from ebaysdk.finding import Connection as Finding


class Workbook():
    
    def __init__(self, filename):
        self.filename = filename
        self.workbook = xlsxwriter.Workbook(self.filename)
        
    def create_worksheet(self, worksheet):
        self.worksheet = self.workbook.add_worksheet(worksheet)
        
    def xlsx_write(self, column_names, r=0, c=0):
        for name in column_names:
            self.worksheet.write(r, c, name)
            c += 1
        
    def close(self):
        self.workbook.close()

def ebay_search(keywords, filters, workbook, tkname):
    
    for keyword in keywords.keys():
        id_dict = {}
        r = 0
        workbook.create_worksheet(keyword)
        api = Finding(appid="********", config_file=None)
        response = api.execute('findItemsAdvanced', {'keywords': keyword}) 
        soup = bs4.BeautifulSoup(response.content, 'lxml')
        items = soup.find_all('item')
    
        for item in items:
            title = item.title.string
            url = item.viewitemurl.string
            item_price = item.currentprice.string
            ship_type = item.shippingtype.string
            list_type = item.listingtype.string
            condition = item.conditiondisplayname.string
            item_id = item.itemid.string
            try:
                ship_price = item.shippingservicecost.string
            except:
                ship_price = 'error'
            try:
                post_code = item.postalcode.string
            except:
                post_code = 'N/A'
            if 'error' in ship_price.lower():
                total_price = float(item_price)
            else:
                total_price = float(item_price) + float(ship_price)
            
            price_limit = keywords[keyword]['price_param']
            stock_limit = keywords[keyword]['stock_param']
            rows = [keyword, title, url, item_price, ship_type, list_type, condition, item_id, ship_price, post_code, total_price]
            row_names = ['keyword', 'title', 'url', 'item_price', 'ship_type', 'list_type', 'condition', 'item_id', 'ship_price', 'post_code', 'total_price']
            
            
            discard = discard_logic(filters, title, total_price, condition, price_limit)
            
            
            if discard == False:
                id_dict[url] = dict(zip(row_names, rows))
        
                        
        row_names.append('stock')
        workbook.xlsx_write(row_names)
        session = requests.Session()
        for url in id_dict.keys():
            stock = stock_search(url, session)
            stock = stock.replace(",", "")
            print(str(id_dict[url]['url']))
            
            print(str(stock))
            if int(stock) >= stock_limit:
                r += 1
                id_dict[url]['stock'] = int(stock)
                rows = []
                for value in id_dict[url].values():
                    rows.append(value)
                workbook.xlsx_write(rows, r)
            tkname.after(random.randint(5000,10000),passit())

    
    print('DONE')
def passit():
    pass
    
def stock_search(url, session):
    headers = {"*******"}
    req = session.get(url, headers=headers)
    soup = bs4.BeautifulSoup(req.text, "html.parser")
    soup = str(soup)
    try:
        stock = soup.split("Please enter a number less than or equal to ")[1]
        stock = stock.split(".")[0]
    except:
        stock = 0   
    return stock
        
            
def discard_logic(filters, title, total_price, condition, price_limit):
    discard = False
    for filter in filters:
        if filter in title.lower():
            discard = True
            break
        if total_price >= float(price_limit):
            discard = True
    if 'new' not in condition.lower():
        discard = True
    return discard
