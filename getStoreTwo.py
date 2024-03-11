import mysql.connector
import requests
import xlsxwriter
from scrapy.selector import Selector
import scrapy



class Search:

    def __init__(self):
        self.row = 1
        self.col = 0
        self.page = 1
        self.headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}
        
 

    def connectDB(self):
        self.mydb = mysql.connector.connect(
        host="localhost",
        user="root",
        password="",
        database="stores"
)
    def setExecl(self):
        # dont stop the bot till the end of the scraping for getting the execl file
        
        self.workbook = xlsxwriter.Workbook('zap12.xlsx')
        self.worksheet = self.workbook.add_worksheet()
        bold = self.workbook.add_format({'bold': True})
        self.worksheet.write('A1', 'websiteName', bold)
        self.worksheet.write('B1', 'WebsiteLink', bold)
        

    def link(self,link):
            page = 0
            while page < 1:
              page = page+1
              self.response = requests.get(link+"&pageinfo="+str(page), headers=self.headers)
              self.getLinks()
              self.getStores()
              self.storeInfo()
            
  
    def getLinks(self):
        self.link = []
        sel = Selector(text=self.response.content)
        head = sel.xpath("//*[@class='withModelRow ModelRowContainer']")
        for heads in head:
           self.link.append(heads.xpath(".//*[@class='ModelTitle']/@href").extract_first())
           print(heads.xpath(".//*[@class='ModelTitle']/@href").extract_first())
       

    def getStores(self):
        self.storeLink = []
        x =''
        print("----")
        for link in self.link:
         
           self.response = requests.get("https://www.zap.co.il"+link, headers=self.headers)
           sel = Selector(text=self.response.content)
           head = sel.xpath("//*[@class='cell2']")
          
          
           for heads in head:
             x = heads.xpath(".//*[@class='compare-item-image']/a/@href").extract_first() 
             if x not in self.storeLink:
                self.storeLink.append(heads.xpath(".//*[@class='compare-item-image']/a/@href").extract_first())
                print(heads.xpath(".//*[@class='compare-item-image']/a/@href").extract_first())
              
             
    
    def storeInfo(self):
        mycursor = self.mydb.cursor(buffered=True)
       
        for link in self.storeLink:
           if link:
            self.response = requests.get("https://www.zap.co.il"+link, headers=self.headers)
            sel = Selector(text=self.response.content)
            head = sel.xpath("//*[@class='StoreInfo']")
         
     
            for heads in head:
             webSitename = heads.xpath(".//h1[@itemprop='name']/a/text()").extract_first()
             webSiteLink =  heads.xpath(".//*[@class='StoreUrl noModelClick']/b/a/text()").extract_first()
             if webSitename is not None and webSiteLink:
                sql = "SELECT * FROM stores WHERE store_link = '%s'" % (webSiteLink)
                mycursor.execute(sql)
                self.worksheet.write(self.row, self.col,webSitename)
                self.worksheet.write(self.row, self.col + 1,webSiteLink)
                self.row += 1
                
                print(webSitename)

                sql = "INSERT INTO stores (store_name, store_link) VALUES (%s, %s)"
                val = (webSitename, webSiteLink)
                mycursor.execute(sql,val)
                self.mydb.commit()
                print(mycursor.rowcount, "record inserted.")
                
       


search = Search()
search.connectDB()
search.setExecl()
search.link("https://www.zap.co.il/models.aspx?sog=c-pclaptop")
search.workbook.close()
