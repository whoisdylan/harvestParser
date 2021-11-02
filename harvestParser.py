from pandas import *

def getDateString(sheetId):
    if len(sheetId) < 2:
        print("unsupported sheetId length: ", len(sheetId), sheetId)
        exit()
    elif len(sheetId) == 2:
        return sheetId[0] + "/" + sheetId[1:2]
    elif len(sheetId) == 3:
        if sheetId[0:2] == "10":
            return sheetId[0:2] + "/" + sheetId[2:3]
        else:
            return sheetId[0] + "/" + sheetId[1:3]
    else:
        if sheetId[0:2] == "10":
            return sheetId[0:2] + "/" + sheetId[2:4]
        else:
            return sheetId[0] + "/" + sheetId[1:3]

def getUniversalCrop(crop):
    cropMap = {
        "Sweet Pepper": "Sweet Peppers",
        "White Icicle Radish": "Radishes",
        "Red Kale": "Kale",
        "Acorn Squash": "Winter Squash (Acorn)",
        "Melon": "Melons",
        "Shelling Bean": "Beans",
        "New Potatoes": "Potatoes",
        "Baby Spinach": "Spinach",
        "Globe Eggplant": "Eggplant",
        "Green Beans": "Beans",
        "Cherry Tomatoes": "Tomatoes (Cherry)",
        "Head Lettuce": "Lettuce Heads",
        "Radish": "Radishes",
        "Winter Radish": "Radishes",
        "Napa Cabbage": "Cabbage",
        "Green Cabbage": "Cabbage",
        "Slicing Tomatoes": "Tomatoes",
        "Turnip": "Turnips",
        "Sweet Corn": "Corn",
        "Curly Kale": "Kale",
        "Snap Peas": "Peas",
        "Bok Choy": "Asian Greens",
        "Lacinato Kale": "Kale",
        "Easter Egg Radish": "Radishes",
        "Cucumber": "Cucumbers",
        "Red Mustard": "Mustard",
        "Green Mustard": "Mustard",
    }

    ahCrops = [
        "Oregano",
        "Mustard Mix",
        "Broccoli Rabe",
        "Scallions",
        "Apples",
        "Artichoke",
        "Arugula",
        "Asian Greens",
        "Baby Bok Choy",
        "Basil",
        "Beans",
        "Beets",
        "Bekana Greens",
        "Bell Peppers",
        "Broccoli",
        "Brussel Sprouts",
        "Bunching Onions",
        "Cabbage",
        "Carrots",
        "Cauliflower",
        "Celery",
        "Celtuce",
        "chicken eggs",
        "Chinese Cabbage",
        "Cilantro",
        "Collards",
        "Corn",
        "Cucumbers",
        "Dill",
        "duck eggs",
        "Dwarf Sunflowers",
        "Eggplant",
        "Fava Beans",
        "Fennel",
        "Flowers",
        "Garlic",
        "Garlic Scapes",
        "Garden Huckleberries",
        "Green Garlic",
        "Ground Cherries",
        "Heirloom Tomatoes",
        "Hot Peppers",
        "Italian Dandelion",
        "Jalapeno Peppers",
        "Jerusalem Artichokes",
        "Kale",
        "Kale Mix",
        "Kohlrabi",
        "Leeks",
        "Lettuce Heads",
        "Lettuce Mix",
        "Lunchbox Peppers",
        "Malabar Spinach",
        "Melons",
        "Mesclun Mix",
        "Mini Eggplant",
        "Mini Lettuce",
        "Mushrooms",
        "Mustard",
        "New Zealand Spinach",
        "Okra",
        "Onions",
        "Parsley",
        "Parsnips",
        "Peas",
        "Potatoes",
        "Pumpkins",
        "Radicchio",
        "Radishes",
        "Salad Mix",
        "Shallots",
        "Shishito Peppers",
        "Spinach",
        "Summer Squash",
        "Sweet Peppers",
        "Sweet Potato",
        "Swiss Chard",
        "Tokyo Bekana",
        "Tomatillos",
        "Tomatoes",
        "Tomatoes (cherry)",
        "Turnips",
        "Watermelon",
        "Winter Squash (Acorn)",
        "Winter Squash (Butternut)",
        "Winter Squash (Delicata)",
        "Winter Squash (Kabocha)",
        "Winter Squash (Spaghetti)",
        "Winter Squash",
    ]
    if crop not in ahCrops and crop not in cropMap:
        print(crop, " not in ahCrops")
    if crop in cropMap:
        return cropMap[crop]
    else:
        return crop
    
def printWeight(date, cropName, totalWeight, destination):
    print(date + ", " + cropName + ", " + totalWeight + ", " + destination)

def getWeightsFromSheet(sheet, sheetId, harvestData):
    if not sheetId[0].isdigit():
        print("skipping ", sheetId)
        return

    # map of destination title in mountain records to destination in our harvest records
    destinationMap = { "Fellowship": "Fellowship", "Large": "CSA", "Small":
    "CSA", "Market": "Farmers Market", "Donation": "Donation" }
    for cropId, crop in sheet['Crop'].items():
        if pandas.isna(crop):
            continue

        totalWeight = round(sheet['Total Weight'][cropId], 2)
        if pandas.isna(totalWeight):
            continue
        
        dateString = getDateString(sheetId) + "/21"
        unitWeight = sheet['Weight (lb)'][cropId]
        cropTotal = 0
        for destinationOld, destinationNew in destinationMap.items():
            if destinationOld not in sheet:
                continue

            unitAmount = sheet[destinationOld][cropId]
            weight = round(unitWeight * unitAmount, 2)
            if weight > 0:
                cropTotal += weight
                universalCrop = getUniversalCrop(crop)
                printWeight(dateString, universalCrop, weight.astype("str"), destinationNew)
                harvestData.append([pandas.to_datetime(dateString).date(), universalCrop, weight, destinationNew])

        # sanity check
        cropTotal = round(cropTotal, 2)
        if totalWeight != cropTotal:
            print("***missing weights: ", sheetId, crop, totalWeight, cropTotal)

def parseFellows():
    workbook = ExcelFile("/Users/dylan/Documents/AF 2021 Fellowship Harvest Tracker.xlsx")
    harvestData = []
    for sheetId in workbook.sheet_names:
        getWeightsFromSheet(workbook.parse(sheetId), sheetId, harvestData)

    pandas.DataFrame(harvestData).to_excel("/Users/dylan/Documents/fellowship_harvest.xlsx", index=False, header=False)

def parseMarkets():
    workbook = ExcelFile("/Users/dylan/Documents/AF 2021 CSA, Market, and Donation Tracker.xlsx")
    harvestData = []
    for sheetId in workbook.sheet_names:
        getWeightsFromSheet(workbook.parse(sheetId), sheetId, harvestData)

    pandas.DataFrame(harvestData).to_excel("/Users/dylan/Documents/markets_harvest.xlsx", index=False, header=False)

if __name__ == "__main__":
    # parseFellows()
    parseMarkets()
