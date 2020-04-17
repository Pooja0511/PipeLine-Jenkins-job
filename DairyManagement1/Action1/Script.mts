'Initialize Plant and Raw Material Information
'Dairy_LIB.material = DataTable("Material","SAP")
'Dairy_LIB.plant = DataTable("Plant","SAP")
RepositoriesCollection.Add(Environment("TestDir") & "\ObjectRepository\SAP_Material_OR.tsr")
RepositoriesCollection.Add(Environment("TestDir") & "\ObjectRepository\SAP_NewBatch_OR.tsr")
RepositoriesCollection.Add(Environment("TestDir") & "\ObjectRepository\SAP_Stock_OR.tsr")
RepositoriesCollection.Add(Environment("TestDir") & "\ObjectRepository\SAP_TimeDepStock_OR.tsr")

ExecuteFile Environment("TestDir") & "\Libs\" & "Sap_LIB.txt"
ExecuteFile Environment("TestDir") & "\Libs\" & "SAP_Material_LIB.txt"
ExecuteFile Environment("TestDir") & "\Libs\" & "SAP_NewBatch_LIB.txt"
ExecuteFile Environment("TestDir") & "\Libs\" & "SAP_Stock_LIB.txt"
ExecuteFile Environment("TestDir") & "\Libs\" & "SAP_TimeDepStock_LIB.txt"
ExecuteFile Environment("TestDir") & "\Libs\" & "UI5_Stock_LIB.txt"
ExecuteFile Environment("TestDir") & "\Libs\" & "Utility_LIB.txt"
ExecuteFile Environment("TestDir") & "\Libs\" & "Web_LIB.txt"
Dim localRegion
localRegion = determineLocalRegionForNumberCalculation()
'UI5_Stock_LIB.localRegionSetting = localRegion
'UI5_Stock_LIB.no_Of_StorageLocations = DataTable("No_Of_SLoc","Global")
'UI5_Stock_LIB.browser_ExeutionMode = DataTable("BrowserName","Global")

'If UI5_Stock_LIB.browser_ExeutionMode = "iexplore" Then
'	RepositoriesCollection.Add(Environment("TestDir") & "\ObjectRepository\UI5_Stock_OR_IE.tsr")
'Else
'	RepositoriesCollection.Add(Environment("TestDir") & "\ObjectRepository\UI5_Stock_OR_Chrome.tsr")
'End If

'******************************** Get actual stock - Start *******************************
checkTransactionStatus(login_SAP ("DT3", 300, "xqt_test", "xqtdairy", "en"))

'Step1- Start Transaction MM03
checkTransactionStatus(startTransaction("MM03"))

'Step2- Get Value in Plan Unit of Measure
SAP_Material_LIB.material = DataTable("Material","SAP")
SAP_Material_LIB.plant = DataTable("Plant","SAP")
checkTransactionStatus(SAP_Material_LIB.getPlantUnitOfMeasure())

'Step3- Verify that Raw Material Relevant Flag is set
checkProperty True,SAP_Material_LIB.raw_Mat_Relevant

'Step4- Get Calculation data for the plan unit of measure.
'checkTransactionStatus(SAP_Material_LIB.getConversionFormula())

'Step5- Start Transaction /nMB5B
checkTransactionStatus(startTransaction("/nMB5B"))

'Step6- Save Stock value for Strorate Location 1001 
SAP_Stock_LIB.material = DataTable("Material","SAP")
SAP_Stock_LIB.plant = DataTable("Plant","SAP")
SAP_Stock_LIB.storageLocation = DataTable("StorageLoc_1001","SAP")
SAP_Stock_LIB.selection_FromDate=Day(Date)&"."&Month(Date)&"."&Year(Date)
SAP_Stock_LIB.selection_ToDate=Day(Date)&"."&Month(Date)&"."&Year(Date)
checkTransactionStatus(SAP_Stock_LIB.getStock())

'Step7- Click Back button
checkTransactionStatus(back_F3())

'Step8- Save Stock value for Strorate Location 1002
SAP_Stock_LIB.storageLocation = DataTable("StorageLoc_1002","SAP")
SAP_Stock_LIB.selection_FromDate=Day(Date)&"."&Month(Date)&"."&Year(Date)
SAP_Stock_LIB.selection_ToDate=Day(Date)&"."&Month(Date)&"."&Year(Date)
checkTransactionStatus(SAP_Stock_LIB.getStock())

'Step8- CLick back button
checkTransactionStatus(back_F3())

'Step9- Save Stock value for Strorate Location 1007
SAP_Stock_LIB.storageLocation = DataTable("StorageLoc_1007","SAP")
SAP_Stock_LIB.selection_FromDate=Day(Date)&"."&Month(Date)&"."&Year(Date)
SAP_Stock_LIB.selection_ToDate=Day(Date)&"."&Month(Date)&"."&Year(Date)
checkTransactionStatus(SAP_Stock_LIB.getStock())

checkTransactionStatus(back_F3())
'Step9A- Save Stock value for Strorate Location 1008
SAP_Stock_LIB.storageLocation = DataTable("StorageLoc_1008","SAP")
SAP_Stock_LIB.selection_FromDate=Day(Date)&"."&Month(Date)&"."&Year(Date)
SAP_Stock_LIB.selection_ToDate=Day(Date)&"."&Month(Date)&"."&Year(Date)
checkTransactionStatus(SAP_Stock_LIB.getStock())

checkTransactionStatus(logout_SAP())
'******************************** Get actual stock - End *******************************


''******************************** Check Opening stock planning without time-dependent stock - Start *******************************
'
''Step10- Login to dairy web portal
'UI5_Stock_LIB.material = DataTable("Material","SAP")
'UI5_Stock_LIB.plant = DataTable("Plant","SAP")
'UI5_Stock_LIB.UserName = DataTable("UserName","Web")
'UI5_Stock_LIB.Password = DataTable("Password","Web")
'checkTransactionStatus(UI5_Stock_LIB.openURL("https://vhostuq3.awscloud.msg.de:4443/sap/bc/ui2/flp?sap-client=300&sap-language=EN#DairyManualOpeningStock-display"))
'
''Step11- Fill in search criteria and check stock opening quantity
''Checkpoint: Quantity RM Raw Milk = Sum of actual stock (step one) 
''Checkpoint: Plan Unit of Measure is the same like in the material master
''Checkpoint: Quantities of components are calculated based on the quantity and the percentage
''Checkpoint: Check the opening stock for each storage location
'UI5_Stock_LIB.x = SAP_Material_LIB.x
'UI5_Stock_LIB.y = SAP_Material_LIB.y
'UI5_Stock_LIB.plan_Unit_Of_Measure = SAP_Material_LIB.plan_Unit_Of_Measure
'checkTransactionStatus(UI5_Stock_LIB.searchDairyStock())
'UI5_Stock_LIB.checkMaterialStock()
'UI5_Stock_LIB.checkOpeningStock_PlantElements()
'
''******************************** Check Opening stock planning without time-dependent stock - End *******************************
'
'
''******************************** Man. Opening Stock - Start *******************************
'
''Step12- If the quanities of manual opening stocks on Storage Location Level are not initial delete every quantity and press save
'checkTransactionStatus(UI5_Stock_LIB.resetOpeningStock())
'
''Step13- Fill in man opening stock with quantity and percentage info
''Checkpoint- The Quantities of the percent are calculated
'UI5_Stock_LIB.quantity = DataTable("RM_Quantity","Web")
'UI5_Stock_LIB.OrganicFatPercent = DataTable("RM_OrganicFatPercent","Web")
'UI5_Stock_LIB.ProteinBioPercent = DataTable("RM_ProteinBioPercent","Web")
'UI5_Stock_LIB.DryMatterPercent = DataTable("RM_DryMatterPercent","Web")
'UI5_Stock_LIB.setOpeningStock()
'
''Step13- Delete Manual Opening Stock and percentages
'checkTransactionStatus(UI5_Stock_LIB.resetOpeningStock())
'
''Step14- Fill in Quanity and Percentage for Storage Location 1001:
''Checkpoint: Quantities on Ingredient Level will be calculated
'UI5_Stock_LIB.quantity = DataTable("StorLoc1001_Quantity","Web")
'UI5_Stock_LIB.OrganicFatPercent = DataTable("StorLoc1001_OrganicFatPercent","Web")
'UI5_Stock_LIB.ProteinBioPercent = DataTable("StorLoc1001_ProteinBioPercent","Web")
'UI5_Stock_LIB.DryMatterPercent = DataTable("StorLoc1001_DryMatterPercent","Web")
'UI5_Stock_LIB.storageLocation = DataTable("StorageLoc_1001","SAP")
'UI5_Stock_LIB.setOpeningStockLoc()
'
''Step15- In Storage Location 1002, only Fill in the Percentages
''Checkpoint: Quantities on Ingredient Level will be calculated
'UI5_Stock_LIB.OrganicFatPercent = DataTable("StorLoc1002_OrganicFatPercent","Web")
'UI5_Stock_LIB.ProteinBioPercent = DataTable("StorLoc1002_ProteinBioPercent","Web")
'UI5_Stock_LIB.DryMatterPercent = DataTable("StorLoc1002_DryMatterPercent","Web")
'UI5_Stock_LIB.storageLocation = DataTable("StorageLoc_1002","SAP")
'UI5_Stock_LIB.setOpeningStockLoc()
'
''Step16- Storage Location 1007, fill in the Quantity
''Checkpoint: Quantities on Ingredient Level will be calculated
'UI5_Stock_LIB.quantity = DataTable("StorLoc1007_Quantity1","Web")
'UI5_Stock_LIB.OrganicFatPercent = DataTable("StorLoc1007_OrganicFatPercent1","Web")
'UI5_Stock_LIB.ProteinBioPercent = DataTable("StorLoc1007_ProteinBioPercent1","Web")
'UI5_Stock_LIB.DryMatterPercent = DataTable("StorLoc1007_DryMatterPercent1","Web")
'UI5_Stock_LIB.storageLocation = DataTable("StorageLoc_1007","SAP")
'UI5_Stock_LIB.setOpeningStockLoc()
'
''Step16A- Storage Location 1008, fill in the Quantity
''Checkpoint: Quantities on Ingredient Level will be calculated
'UI5_Stock_LIB.quantity = DataTable("StorLoc1008_Quantity","Web")
'UI5_Stock_LIB.OrganicFatPercent = DataTable("StorLoc1008_OrganicFatPercent","Web")
'UI5_Stock_LIB.ProteinBioPercent = DataTable("StorLoc1008_ProteinBioPercent","Web")
'UI5_Stock_LIB.DryMatterPercent = DataTable("StorLoc1008_DryMatterPercent","Web")
'UI5_Stock_LIB.storageLocation = DataTable("StorageLoc_1008","SAP")
'UI5_Stock_LIB.setOpeningStockLoc()
'
''Step17- Check Quanitites on Plant level:
''Checkpoint: The Quantites on Plant level are the Sum of the quantities on storage location Level.
'UI5_Stock_LIB.checkManualStock()
''Reset Browser zoom level
'fnResetBrowserZoomLevel()
'
'closeBrowser(UI5_Stock_LIB.browser_ExeutionMode)
'
''******************************** Man. Opening Stock - End *******************************
'
''******************************** Add new batch - Start *****************************
'checkTransactionStatus(login_SAP ("DT3", 300, "xqt_test", "xqtdairy", "en"))
'
''Step18- Start Transaction MSC1n
'startTransaction("/n")
'startTransaction("MSC1n")
'
''Step19- Create a new batch and save batch number
'SAP_NewBatch_LIB.material = DataTable("Material","SAP")
'SAP_NewBatch_LIB.plant = DataTable("Plant","SAP")
'SAP_NewBatch_LIB.storageLocation = DataTable("StorageLoc_1001","SAP")
'SAP_NewBatch_LIB.fatContent = DataTable("fatContent","SAP")
'SAP_NewBatch_LIB.proteinContent = DataTable("proteinContent","SAP")
'SAP_NewBatch_LIB.dryMatterContent = DataTable("dryMatterContent","SAP")
'SAP_NewBatch_LIB.timeDept_flag = UI5_Stock_LIB.timeDept_flag
'SAP_NewBatch_LIB.plan_Unit_Of_Measure = SAP_Material_LIB.plan_Unit_Of_Measure
'checkTransactionStatus(SAP_NewBatch_LIB.newBatch())
'
''Step20- Save
'checkTransactionStatus(save())
'
''Step21- Checkpoint: Successmessage in status bar
'checkStatusBarMessage("Creating batch.*")
'
''Step21- Start Transaction MI10
'checkTransactionStatus(startTransaction("/n"))
'checkTransactionStatus(startTransaction("MI10"))
'
''Step22- Create Inventury
'SAP_NewBatch_LIB.countDate = Day(Date)&"."&Month(Date)&"."&Year(Date)
'SAP_NewBatch_LIB.documentDate = Day(Date)&"."&Month(Date)&"."&Year(Date)
'SAP_NewBatch_LIB.storageLocation = DataTable("StorageLoc_1001","SAP")
'SAP_NewBatch_LIB.quantity = DataTable("quantity","SAP")
'checkTransactionStatus(SAP_NewBatch_LIB.createInventury())
'
''Step23- Save
'checkTransactionStatus(save())
'
''Step24- Check successful status message for Inventury creation
'If(SAP_NewBatch_LIB.timeDept_flag) Then
'	checkStatusBarMessage("Phys. inventory document.* posted without differences")
'Else
'	checkStatusBarMessage("Diffs in phys. inv. doc. .* posted with m. doc. .*")
'End If
'
''******************************** Add new batch - End *****************************
'
'
''******************************** Run Report for time-dependent stock - Start *****************************
'
''Step25- Start Transaction SA38
'checkTransactionStatus(startTransaction("/n"))
'checkTransactionStatus(startTransaction("SA38"))
'
''Step26- Create Time dependent Stock Report
''Check successful message “Processing raw material D10001”
'SAP_TimeDepStock_LIB.rawMaterial = DataTable("Material","SAP")
'SAP_TimeDepStock_LIB.plant = DataTable("Plant","SAP")
'SAP_TimeDepStock_LIB.program = DataTable("program","SAP")
'checkTransactionStatus(SAP_TimeDepStock_LIB.refreshTimeDepStock())
'
'checkTransactionStatus(logout_SAP())
'
''******************************** Run Report for time-dependent stock - End *****************************
'
'
''******************************** Opening & Timedependent stock planning - Start *****************************
'checkTransactionStatus(UI5_Stock_LIB.openURL("https://vhostuq3.awscloud.msg.de:4443/sap/bc/ui2/flp?sap-client=300&sap-language=EN#DairyManualOpeningStock-display"))
'
''Step27- Fill in search criteria, display the all stock values
''Opening Stock: 
''Checkpoint: Quantity RM Raw Milk = Sum of actual stock (step one) 
''Checkpoint: Quantities of components are calculated based on the quantity and the percentage
''refreshBrowser()
'checkTransactionStatus(UI5_Stock_LIB.searchDairyStock())
''UI5_Stock_LIB.checkMaterialStock()
'
''Time Dependent Stock: 
''Checkpoint: Time-Dependent stock = Quanity + 1000
''Checkpoint Fat Quantity: Quantity of the component = Quantity of the component in the opening stock + 40 (according to the characteristic in the batch. Better use variable)
''Checkpoint Fat Percentage = Quantity of the component * 100 / Quantity Time-Dependent Stock
''Checkpoint Protein  Quantity: Quantity of the component = Quantity of the component in the opening stock + 36 (according to the characteristic in the batch. Better use variable)
''Checkpoint Protein Percentage = Quantity of the component * 100 / Quantity Time-Dependent Stock
''Checkpoint Dry Matter  Quantity: Quantity of the component = Quantity of the component in the opening stock + 120 (according to the characteristic in the batch. Better use variable)
''Checkpoint Dry Matter Percentage = Quantity of the component * 100 / Quantity Time-Dependent Stock
''Check the opening stock for each storage location. The Time-Dependent Stock in storage location 1001 is 1000L higher than the opening stock. The other fields are the same
'UI5_Stock_LIB.checkTimeDepMaterialStock()
'
''Step28- If the quanities of manual opening stocks on Storage Location Level are not initial delete every quantity and press save
'checkTransactionStatus(UI5_Stock_LIB.resetOpeningStock())
'
''Step29- Fill in Man. Opening Stock quantity and percentage at plant level
''Checkpoint: The Quantities of the percent are calculated:
'UI5_Stock_LIB.quantity = DataTable("RM_Quantity","Web")
'UI5_Stock_LIB.OrganicFatPercent = DataTable("RM_OrganicFatPercent","Web")
'UI5_Stock_LIB.ProteinBioPercent = DataTable("RM_ProteinBioPercent","Web")
'UI5_Stock_LIB.DryMatterPercent = DataTable("RM_DryMatterPercent","Web")
'UI5_Stock_LIB.setOpeningStock()
'
''Step30- Reset man opening stock
'checkTransactionStatus(UI5_Stock_LIB.resetOpeningStock())
'
''Step31- Fill in Quanitys and Percentages for Storage Location 1001:
''Checkpoint: Quantities on Ingredient Level will be calculated
'UI5_Stock_LIB.quantity = DataTable("StorLoc1001_Quantity","Web")
'UI5_Stock_LIB.OrganicFatPercent = DataTable("StorLoc1001_OrganicFatPercent","Web")
'UI5_Stock_LIB.ProteinBioPercent = DataTable("StorLoc1001_ProteinBioPercent","Web")
'UI5_Stock_LIB.DryMatterPercent = DataTable("StorLoc1001_DryMatterPercent","Web")
'UI5_Stock_LIB.storageLocation = DataTable("StorageLoc_1001","SAP")
'UI5_Stock_LIB.setOpeningStockLoc()
'
''Step32- Storage Location 1002 only Fill in the Percentages
'UI5_Stock_LIB.OrganicFatPercent = DataTable("StorLoc1002_OrganicFatPercent","Web")
'UI5_Stock_LIB.ProteinBioPercent = DataTable("StorLoc1002_ProteinBioPercent","Web")
'UI5_Stock_LIB.DryMatterPercent = DataTable("StorLoc1002_DryMatterPercent","Web")
'UI5_Stock_LIB.storageLocation = DataTable("StorageLoc_1002","SAP")
'UI5_Stock_LIB.setOpeningStockLoc()
'
''Step33- Change Quanitity for the material on plant level 
''Checkpoint: Error Message
''Checkpoint: the data can not be saved 
'UI5_Stock_LIB.quantity = "100.000,000"
'UI5_Stock_LIB.changePlantQuantity()
'
''Step34- Reset man opening stock
'checkTransactionStatus(UI5_Stock_LIB.resetOpeningStock())
'
''Step 35- Fill in Quanitys and Percentages for Storage Location 1001:
''Checkpoint: Quantities on Ingredient Level will be calculated
'UI5_Stock_LIB.quantity = DataTable("StorLoc1001_Quantity","Web")
'UI5_Stock_LIB.OrganicFatPercent = DataTable("StorLoc1001_OrganicFatPercent","Web")
'UI5_Stock_LIB.ProteinBioPercent = DataTable("StorLoc1001_ProteinBioPercent","Web")
'UI5_Stock_LIB.DryMatterPercent = DataTable("StorLoc1001_DryMatterPercent","Web")
'UI5_Stock_LIB.storageLocation = DataTable("StorageLoc_1001","SAP")
'UI5_Stock_LIB.setOpeningStockLoc()
'
''Step 36- In Storage Location 1002 Only Fill in the Percentages:
''Checkpoint: Quantities on Ingredient Level will be calculated
'UI5_Stock_LIB.OrganicFatPercent = DataTable("StorLoc1002_OrganicFatPercent","Web")
'UI5_Stock_LIB.ProteinBioPercent = DataTable("StorLoc1002_ProteinBioPercent","Web")
'UI5_Stock_LIB.DryMatterPercent = DataTable("StorLoc1002_DryMatterPercent","Web")
'UI5_Stock_LIB.storageLocation = DataTable("StorageLoc_1002","SAP")
'UI5_Stock_LIB.setOpeningStockLoc()
'
''Step 37- In Storage Location 1007 input quantity:
''Checkpoint: Quantities on Ingredient Level will be calculated
'UI5_Stock_LIB.quantity = DataTable("StorLoc1007_Quantity2","Web")
'UI5_Stock_LIB.OrganicFatPercent = DataTable("StorLoc1007_OrganicFatPercent2","Web")
'UI5_Stock_LIB.ProteinBioPercent = DataTable("StorLoc1007_ProteinBioPercent2","Web")
'UI5_Stock_LIB.DryMatterPercent = DataTable("StorLoc1007_DryMatterPercent2","Web")
'UI5_Stock_LIB.storageLocation = DataTable("StorageLoc_1007","SAP")
'UI5_Stock_LIB.setOpeningStockLoc()
'
''Step37A- Storage Location 1008, fill in the Quantity
''Checkpoint: Quantities on Ingredient Level will be calculated
'UI5_Stock_LIB.quantity = DataTable("StorLoc1008_Quantity","Web")
'UI5_Stock_LIB.OrganicFatPercent = DataTable("StorLoc1008_OrganicFatPercent","Web")
'UI5_Stock_LIB.ProteinBioPercent = DataTable("StorLoc1008_ProteinBioPercent","Web")
'UI5_Stock_LIB.DryMatterPercent = DataTable("StorLoc1008_DryMatterPercent","Web")
'UI5_Stock_LIB.storageLocation = DataTable("StorageLoc_1008","SAP")
'UI5_Stock_LIB.setOpeningStockLoc()
'
''Step 38- Check Quanitites on Plant level:
''Checkpoint: The Quantities on Plant level are the Sum of the quantities on storage location Level. (Also the quantities of the components) 
''Checkpoint Protein Percentage = Quantity of the component * 100 / Quantity man. Opening Stock
'UI5_Stock_LIB.checkManualStock()
'Browser("Plan Opening Stocks_2").Page("Plan Opening Stocks").SAPUIButton("Save").Click
'
''Reset Browser zoom level
'fnResetBrowserZoomLevel()
'closeBrowser(UI5_Stock_LIB.browser_ExeutionMode)
