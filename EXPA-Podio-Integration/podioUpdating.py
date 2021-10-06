from pypodio2 import api

podio_client_id = ''   #write the client id
podio_client_secret = ''   #write the client secret
podio_login_id = ''   #write the username
podio_login_pw = ''   #write the password
podio = api.OAuthClient(podio_client_id, podio_client_secret, podio_login_id, podio_login_pw)



def podioUpdateAPDsICX(data, appID):
    newAddedItems = 0
    notAddedItems = 0
    if len(data) != 0:
        duplicatedItems, notDuplicatedItems = checkDuplications(data, appID)
        for i in range(len(notDuplicatedItems)):
            if notDuplicatedItems[i][4] == "IGV":
                productID = 1
            elif notDuplicatedItems[i][4] == "IGTa":
                productID = 2
            elif notDuplicatedItems[i][4] == "IGTe":
                productID = 3
            elif notDuplicatedItems[i][4] == "OGV":
                productID = 4
            elif notDuplicatedItems[i][4] == "OGTa":
                productID = 5
            elif notDuplicatedItems[i][4] == "OGTe":
                productID = 6
            form_item = {
                "fields": [
                    {"external_id": "title", "values": [{"value": notDuplicatedItems[i][0]}]},
                    {"external_id": "ep-name", "values": [{"value": notDuplicatedItems[i][1]}]},
                    {"external_id": "ep-profile-link", "values": [{"url": notDuplicatedItems[i][2]}]},
                    {"external_id": "email", "values": [{"value": notDuplicatedItems[i][3]}]},
                    {"external_id": "product", "values": [{"value": productID}]},
                    {"external_id": "sending-lc", "values": [{"value": notDuplicatedItems[i][5]}]},
                    {"external_id": "sending-mc", "values": [{"value": notDuplicatedItems[i][6]}]},
                    {"external_id": "hosting-lc", "values": [{"value": notDuplicatedItems[i][7]}]},
                    {"external_id": "hosting-mc", "values": [{"value": notDuplicatedItems[i][8]}]},
                    {"external_id": "opportunity-link", "values": [{"url": notDuplicatedItems[i][9]}]},
                    {"external_id": "approval-date", "values": [{"start": notDuplicatedItems[i][10] + " 00:00:00"}]},
                    {"external_id": "slot-start-date", "values": [{"start": notDuplicatedItems[i][11] + " 00:00:00"}]},
                    {"external_id": "slot-end-date", "values": [{"start": notDuplicatedItems[i][12] + " 00:00:00"}]},
                ]
            }
            podio.Item.create(appID, form_item)
            newAddedItems = newAddedItems + 1
            notAddedItems = len(list(duplicatedItems.keys()))
        return newAddedItems, notAddedItems
    return newAddedItems, notAddedItems


def podioUpdateREsICX(data, appID):
    newAddedItems, UpdatedItems = 0, 0
    if len(data) != 0:
        duplicatedItems, notDuplicatedItems = checkDuplications(data, appID)
        for i in range(len(notDuplicatedItems)):
            if notDuplicatedItems[i][4] == "IGV":
                productID = 1
            elif notDuplicatedItems[i][4] == "IGTa":
                productID = 2
            elif notDuplicatedItems[i][4] == "IGTe":
                productID = 3
            elif notDuplicatedItems[i][4] == "OGV":
                productID = 4
            elif notDuplicatedItems[i][4] == "OGTa":
                productID = 5
            elif notDuplicatedItems[i][4] == "OGTe":
                productID = 6
            if notDuplicatedItems[i][11] is None:
                slotStartDate = "2021-08-01"
                slotEndDate = "2021-09-12"
            else:
                slotStartDate = notDuplicatedItems[i][11]
                slotEndDate = notDuplicatedItems[i][12]
            if notDuplicatedItems[i][13] is None:
                reDate = "2021-08-01"
            else:
                reDate = notDuplicatedItems[i][13]
            if notDuplicatedItems[i][14] is None:
                fiDate ="2021-09-12"
            else:
                fiDate = notDuplicatedItems[i][14]
            form_item = {
                "fields": [
                    {"external_id": "title", "values": [{"value": notDuplicatedItems[i][0]}]},
                    {"external_id": "ep-name", "values": [{"value": notDuplicatedItems[i][1]}]},
                    {"external_id": "ep-profile-link", "values": [{"url": notDuplicatedItems[i][2]}]},
                    {"external_id": "email", "values": [{"value": notDuplicatedItems[i][3]}]},
                    {"external_id": "product", "values": [{"value": productID}]},
                    {"external_id": "sending-lc", "values": [{"value": notDuplicatedItems[i][5]}]},
                    {"external_id": "sending-mc", "values": [{"value": notDuplicatedItems[i][6]}]},
                    {"external_id": "hosting-lc", "values": [{"value": notDuplicatedItems[i][7]}]},
                    {"external_id": "hosting-mc", "values": [{"value": notDuplicatedItems[i][8]}]},
                    {"external_id": "opportunity-link", "values": [{"url": notDuplicatedItems[i][9]}]},
                    {"external_id": "approval-date", "values": [{"start": notDuplicatedItems[i][10] + " 00:00:00"}]},
                    {"external_id": "slot-start-date","values": [{"start": slotStartDate + " 00:00:00"}]},
                    {"external_id": "slot-end-date", "values": [{"start": slotEndDate + " 00:00:00"}]},
                    {"external_id": "realization-date","values": [{"start": reDate + " 00:00:00"}]},
                    {"external_id": "experience-end-dates", "values": [{"start": fiDate + " 00:00:00"}]}
                ]
                }
            podio.Item.create(appID, form_item)
            newAddedItems = newAddedItems + 1
        values = list(duplicatedItems.values())
        keys = list(duplicatedItems.keys())
        i = 0
        for val in values:
            form_item = {
                "fields": [
                    {"external_id": "realization-date","values": [{"start": val[13] + " 00:00:00"}]},
                    {"external_id": "experience-end-dates", "values": [{"start": val[14] + " 00:00:00"}]}
                    ]
                }
            podio.Item.update(keys[i], form_item)
            i = i + 1
            UpdatedItems = UpdatedItems + 1
        return newAddedItems,UpdatedItems
    return newAddedItems, UpdatedItems


def podioUpdateAPDsOGX(data, appID):
    newAddedItems = 0
    notAddedItems = 0
    if len(data) != 0:
        duplicatedItems, notDuplicatedItems = checkDuplications(data, appID)
        for i in range(len(notDuplicatedItems)):
            if notDuplicatedItems[i][4] == "IGV":
                productID = 1
            elif notDuplicatedItems[i][4] == "IGTa":
                productID = 2
            elif notDuplicatedItems[i][4] == "IGTe":
                productID = 3
            elif notDuplicatedItems[i][4] == "OGV":
                productID = 1
            elif notDuplicatedItems[i][4] == "OGTa":
                productID = 2
            elif notDuplicatedItems[i][4] == "OGTe":
                productID = 3
            form_item = {
                "fields": [
                    {"external_id": "title", "values": [{"value": notDuplicatedItems[i][0]}]},
                    {"external_id": "ep-name", "values": [{"value": notDuplicatedItems[i][1]}]},
                    {"external_id": "ep-profile-link", "values": [{"url": notDuplicatedItems[i][2]}]},
                    {"external_id": "email", "values": [{"value": notDuplicatedItems[i][3]}]},
                    {"external_id": "product", "values": [{"value": productID}]},
                    {"external_id": "sending-lc", "values": [{"value": notDuplicatedItems[i][5]}]},
                    {"external_id": "sending-mc", "values": [{"value": notDuplicatedItems[i][6]}]},
                    {"external_id": "hosting-lc", "values": [{"value": notDuplicatedItems[i][7]}]},
                    {"external_id": "hosting-mc", "values": [{"value": notDuplicatedItems[i][8]}]},
                    {"external_id": "opportunity-link", "values": [{"url": notDuplicatedItems[i][9]}]},
                    {"external_id": "approval-date", "values": [{"start": notDuplicatedItems[i][10] + " 00:00:00"}]},
                    {"external_id": "slot-start-date", "values": [{"start": notDuplicatedItems[i][11]+ " 00:00:00"}]},
                    {"external_id": "slot-end-date", "values": [{"start": notDuplicatedItems[i][12] + " 00:00:00"}]},
                ]
            }
            podio.Item.create(appID, form_item)
            newAddedItems = newAddedItems + 1
            notAddedItems = len(list(duplicatedItems.keys()))
        return newAddedItems, notAddedItems
    return newAddedItems, notAddedItems


def podioUpdateREsOGX(data, appID):
    newAddedItems, UpdatedItems = 0, 0
    if len(data) != 0:
        duplicatedItems, notDuplicatedItems = checkDuplications(data, appID)
        for i in range(len(notDuplicatedItems)):
            if notDuplicatedItems[i][4] == "IGV":
                productID = 1
            elif notDuplicatedItems[i][4] == "IGTa":
                productID = 2
            elif notDuplicatedItems[i][4] == "IGTe":
                productID = 3
            elif notDuplicatedItems[i][4] == "OGV":
                productID = 1
            elif notDuplicatedItems[i][4] == "OGTa":
                productID = 2
            elif notDuplicatedItems[i][4] == "OGTe":
                productID = 3
            if notDuplicatedItems[i][11] is None:
                slotStartDate = "0000-00-00"
                slotEndDate = "0000-00-00"
            else:
                slotStartDate = notDuplicatedItems[i][11]
                slotEndDate = notDuplicatedItems[i][12]
            form_item = {
                "fields": [
                    {"external_id": "title", "values": [{"value": notDuplicatedItems[i][0]}]},
                    {"external_id": "ep-name", "values": [{"value": notDuplicatedItems[i][1]}]},
                    {"external_id": "ep-profile-link", "values": [{"url": notDuplicatedItems[i][2]}]},
                    {"external_id": "email", "values": [{"value": notDuplicatedItems[i][3]}]},
                    {"external_id": "product", "values": [{"value": productID}]},
                    {"external_id": "sending-lc", "values": [{"value": notDuplicatedItems[i][5]}]},
                    {"external_id": "sending-mc", "values": [{"value": notDuplicatedItems[i][6]}]},
                    {"external_id": "hosting-lc", "values": [{"value": notDuplicatedItems[i][7]}]},
                    {"external_id": "hosting-mc", "values": [{"value": notDuplicatedItems[i][8]}]},
                    {"external_id": "opportunity-link", "values": [{"url": notDuplicatedItems[i][9]}]},
                    {"external_id": "approval-date", "values": [{"start": notDuplicatedItems[i][10] + " 00:00:00"}]},
                    {"external_id": "slot-start-date","values": [{"start": slotStartDate + " 00:00:00"}]},
                    {"external_id": "slot-end-date", "values": [{"start": slotEndDate + " 00:00:00"}]},
                    {"external_id": "realization-date","values": [{"start": notDuplicatedItems[i][13] + " 00:00:00"}]},
                    {"external_id": "experience-end-dates", "values": [{"start": notDuplicatedItems[i][14] + " 00:00:00"}]}
                ]
                }
            podio.Item.create(appID, form_item)
            newAddedItems = newAddedItems + 1
        values = list(duplicatedItems.values())
        keys = list(duplicatedItems.keys())
        i = 0
        for val in values:
            form_item = {
                "fields": [
                    {"external_id": "realization-date","values": [{"start": val[13] + " 00:00:00"}]},
                    {"external_id": "experience-end-dates", "values": [{"start": val[14] + " 00:00:00"}]}
                    ]
                }
            podio.Item.update(keys[i], form_item)
            i = i + 1
            UpdatedItems = UpdatedItems + 1
        return newAddedItems,UpdatedItems
    return newAddedItems, UpdatedItems


def checkDuplications(data, appID):
    appItems = podio.Item.filter(appID, attributes={"limit": 500, "offset": 0})
    duplicatedItems = {}
    notDuplicatedItems = []
    for item in data:
        duplicated = False
        for appItem in appItems['items']:
            if appItem['title'] == item[0]:
                duplicatedItems[appItem['item_id']] = item
                duplicated = True
        if not duplicated:
            notDuplicatedItems.append(item)
    return duplicatedItems, notDuplicatedItems


def printValues(data, newAddedItems, UpdatedItems, dataType):
    if dataType == "OGX APDs" or dataType == "ICX APDs":
        print("newAddedItems" + dataType + ": " + str(newAddedItems))
        print("notAddedItems" + dataType + ": " + str(UpdatedItems))
    else:
        print("newAddedItems" + dataType + ": " + str(newAddedItems))
        print("UpdatedItems" + dataType + ": " + str(UpdatedItems))
