import requests
import xlsxwriter
import sys

# Configuration variables
baseAPI = 'https://prices.azure.com/api/retail/prices?$filter=serviceFamily eq '
listOfAttributes = ['serviceFamily', 'serviceName', 'serviceId', 'productName', 'productId', 'meterName',
                    'unitOfMeasure', 'location', 'unitPrice', 'retailPrice']
listOfServiceFamilies = ["'Compute'", "'Databases'", "'Blockchain'"]
additionalFilter = "and location eq 'EU West'"  # example form: and location eq 'EU West'
sys.setrecursionlimit(1500) # careful when increasing, unset on low-RAM PC

def requestAndParse(url):
    response = requests.get(url)
    response = response.json()

    for item in response["Items"]:
        line = [None] * len(listOfAttributes)

        for key, value in item.items():
            if key in listOfAttributes:
                line[listOfAttributes.index(key)] = value

        allServiceProducts.append(line)

    # check if there is another page of products (because REST response max size 100 items)
    if not response["NextPageLink"]:
        return
    else:
        print("More items - requesting ", response["NextPageLink"])
        requestAndParse(response["NextPageLink"])


def writeToExcel(title, dataList):
    worksheet = workbook.add_worksheet(title.replace("'", " "))
    for col_num, data in enumerate(listOfAttributes):
        worksheet.write(0, col_num, data, headerFormat)
    for row in range(len(dataList)):
        for col in range(len(listOfAttributes)):
            worksheet.write(row + 1, col, dataList[row][col])
    worksheet.autofilter(0, 0, 0, len(listOfAttributes) - 1)
    worksheet.set_column(0, len(listOfAttributes) - 1, 15)


############################ main ############################

workbook = xlsxwriter.Workbook('allDatabaseProducts.xlsx')
headerFormat = workbook.add_format({'bold': True})
headerFormat.set_bg_color('#D9D9D9')
headerFormat.set_border()

# get all products for each specified service plus optional filters and write the content into an excel worksheet
for serviceFamily in listOfServiceFamilies:
    allServiceProducts = []
    print("Fetching products for service: ", serviceFamily, " with filter ", additionalFilter)
    requestAndParse(baseAPI + serviceFamily + additionalFilter)
    print("Done fetching products for service ", serviceFamily)
    print("Found {} products".format(len(allServiceProducts)))
    writeToExcel(serviceFamily, allServiceProducts)

workbook.close()
