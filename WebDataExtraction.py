import pandas as pd
from selenium import webdriver
from bs4 import BeautifulSoup
import mysql.connector as MySQLdb

try:
    driver = webdriver.Chrome("C:/Users/Narasimha/Documents/Python Scripts/Custom development/Drivers/chromedriver")

    products=[] #List to store name of the product
    prices=[] #List to store price of the product
    ratings=[] #List to store rating of the product
    driver.maximize_window()
    driver.get("https://www.flipkart.com/laptops/~buyback-guarantee-on-laptops-/pr?sid=6bo%2Cb5g&uniq")
    content = driver.page_source
    soup = BeautifulSoup(content,features="html.parser")
    for a in soup.findAll('a',href=True, attrs={'class':'_31qSD5'}):
        name=a.find('div', attrs={'class':'_3wU53n'})
        price=a.find('div', attrs={'class':'_1vC4OE _2rQ-NK'})
        rating=a.find('div', attrs={'class':'hGSR34'})
        products.append(name.text)
        prices.append(price.text)
        ratings.append(rating.text)

    # Establish a MySQL connection
    database = MySQLdb.connect (host='localhost', user = 'root', passwd = 'July_1995@', db = 'developement')
    # Get the cursor, which is used to traverse the database, line by line
    cursor = database.cursor(buffered=True)
    #get column names
    cursor.execute('select * from product_data')
    names = tuple(map(lambda x: x[0], cursor.description))
    #create placeholder for values
    placeHolder = '%s,' * len(names)
    placeHolder = placeHolder[0:len(placeHolder)-1]
    #convert column names to string
    temp=''
    for _i in names:
        temp=",".join(names)

    query = 'INSERT INTO product_data ({}) VALUES ({})'.format(temp,placeHolder)

    for i in range(len(products)-1):
        values = []
        var1,var2,var3=products[i],prices[i],ratings[i]
        values.append(var1)
        values.append(var2)
        values.append(var3)
        values = tuple(values)
        cursor.execute(query, values)

    cursor.close()
    database.commit()
    database.close()

    # df = pd.DataFrame({'Product Name':products,'Price':prices,'Rating':ratings})
    # df.to_excel('products.xlsx', index=False, encoding='utf-8')

finally:
    driver.close()