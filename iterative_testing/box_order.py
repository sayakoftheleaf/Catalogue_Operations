def mapBoxes(boxInformationOrder):
   
    # Removing spaces and converting everything to uppercase
    boxInformationOrder = boxInformationOrder.replace(" ", "")
    boxInformationOrder = boxInformationOrder.replace(",", "")
    boxInformationOrder = boxInformationOrder.upper()

    orderDict = {}

    for value in ['L', 'B', 'H', 'W']:
        orderDict[value] = boxInformationOrder.index(value)

    print(orderDict)
    return orderDict

# if (_name_ == _main_):
#    mapBoxes()
