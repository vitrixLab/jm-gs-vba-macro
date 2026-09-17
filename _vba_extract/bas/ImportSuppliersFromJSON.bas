Attribute VB_Name = "ImportSuppliersFromJSON"

Option Explicit



Sub ImportSuppliersFromJSON()

    Dim ws As Worksheet

    Dim rowNum As Long

    Dim i As Long

    Dim lastRow As Long

    

    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")

    

    ' Clear old data (keep headers)

    ws.Rows("2:" & ws.Rows.count).Clear

    

    rowNum = 2

    

    ' ============================================================

    ' COLUMNS: B=NO, C=TIN#, D=PARTICULARS, E=ADDRESS, F=CHART OF ACCOUNT

    ' ============================================================

    

    ' 1

    ws.Cells(rowNum, "B").value = 1

    ws.Cells(rowNum, "C").value = "008-888-659"

    ws.Cells(rowNum, "D").value = "VPHARMA HEALTH AND WELNESS INC"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Clinic Materials and Supplies"

    rowNum = rowNum + 1

    

    ' 2

    ws.Cells(rowNum, "B").value = 2

    ws.Cells(rowNum, "C").value = "246-099-058"

    ws.Cells(rowNum, "D").value = "TRAVELLERS INTL HOTEL INC"

    ws.Cells(rowNum, "E").value = "NEWPORT, PASAY"

    ws.Cells(rowNum, "F").value = "Transportation and Travel"

    rowNum = rowNum + 1

    

    ' 3

    ws.Cells(rowNum, "B").value = 3

    ws.Cells(rowNum, "C").value = "008-120-744"

    ws.Cells(rowNum, "D").value = "RUSTAN"

    ws.Cells(rowNum, "E").value = "SAN LORENZO, MAKATI"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 4

    ws.Cells(rowNum, "B").value = 4

    ws.Cells(rowNum, "C").value = "008-625-341"

    ws.Cells(rowNum, "D").value = "THE MESSHALL"

    ws.Cells(rowNum, "E").value = "MAGALLANES, MAKATI"

    ws.Cells(rowNum, "F").value = "Miscellaneous"

    rowNum = rowNum + 1

    

    ' 5

    ws.Cells(rowNum, "B").value = 5

    ws.Cells(rowNum, "C").value = "252-126-611"

    ws.Cells(rowNum, "D").value = "EUROFUEL INC"

    ws.Cells(rowNum, "E").value = "183 PASAY"

    ws.Cells(rowNum, "F").value = "Fuel and Oil"

    rowNum = rowNum + 1

    

    ' 6

    ws.Cells(rowNum, "B").value = 6

    ws.Cells(rowNum, "C").value = "008-909-992"

    ws.Cells(rowNum, "D").value = "SOUTHEASTASIA RETAIL INC"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 7

    ws.Cells(rowNum, "B").value = 7

    ws.Cells(rowNum, "C").value = "000-388-771"

    ws.Cells(rowNum, "D").value = "JOLLIBEE 39TH"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 8

    ws.Cells(rowNum, "B").value = 8

    ws.Cells(rowNum, "C").value = "207-961-175"

    ws.Cells(rowNum, "D").value = "SAVEMORE"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 9

    ws.Cells(rowNum, "B").value = 9

    ws.Cells(rowNum, "C").value = "001-799-123"

    ws.Cells(rowNum, "D").value = "UNIGLOBE TRAVELWARE INC"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Transportation and Travel"

    rowNum = rowNum + 1

    

    ' 10

    ws.Cells(rowNum, "B").value = 10

    ws.Cells(rowNum, "C").value = "148-855-506"

    ws.Cells(rowNum, "D").value = "MINDWERKS"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 11

    ws.Cells(rowNum, "B").value = 11

    ws.Cells(rowNum, "C").value = "007-851-409"

    ws.Cells(rowNum, "D").value = "PRIMERA CLASE GAS STATION"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Fuel and Oil"

    rowNum = rowNum + 1

    

    ' 12

    ws.Cells(rowNum, "B").value = 12

    ws.Cells(rowNum, "C").value = "200-035-311"

    ws.Cells(rowNum, "D").value = "ACE HARDWARE"

    ws.Cells(rowNum, "E").value = "ALMANZA UNO, LAS PIÑAS"

    ws.Cells(rowNum, "F").value = "Repairs and Maintenance"

    rowNum = rowNum + 1

    

    ' 13

    ws.Cells(rowNum, "B").value = 13

    ws.Cells(rowNum, "C").value = "000-388-771"

    ws.Cells(rowNum, "D").value = "JOLLIBEE BLUEBAY WALK"

    ws.Cells(rowNum, "E").value = "76 PASAY"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 14

    ws.Cells(rowNum, "B").value = 14

    ws.Cells(rowNum, "C").value = "246-969-491"

    ws.Cells(rowNum, "D").value = "S AND R"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 15

    ws.Cells(rowNum, "B").value = 15

    ws.Cells(rowNum, "C").value = "648-700-952"

    ws.Cells(rowNum, "D").value = "GOOD MORNING FOOD"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 16

    ws.Cells(rowNum, "B").value = 16

    ws.Cells(rowNum, "C").value = "008-087-158"

    ws.Cells(rowNum, "D").value = "PNC GARBROS FOODS INC"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 17

    ws.Cells(rowNum, "B").value = 17

    ws.Cells(rowNum, "C").value = "208-006-656"

    ws.Cells(rowNum, "D").value = "SM STORE"

    ws.Cells(rowNum, "E").value = "ORANBO, PASIG"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 18

    ws.Cells(rowNum, "B").value = 18

    ws.Cells(rowNum, "C").value = "214-706-591"

    ws.Cells(rowNum, "D").value = "WATSONS"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 19

    ws.Cells(rowNum, "B").value = 19

    ws.Cells(rowNum, "C").value = "608-395-184"

    ws.Cells(rowNum, "D").value = "GIANETO PIZZERIA"

    ws.Cells(rowNum, "E").value = "ORANBO, PASIG"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 20

    ws.Cells(rowNum, "B").value = 20

    ws.Cells(rowNum, "C").value = "010-624-401"

    ws.Cells(rowNum, "D").value = "MIT SUKOSHI FRESH"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 21

    ws.Cells(rowNum, "B").value = 21

    ws.Cells(rowNum, "C").value = "655-591-162"

    ws.Cells(rowNum, "D").value = "MINISO"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 22

    ws.Cells(rowNum, "B").value = 22

    ws.Cells(rowNum, "C").value = "000-405-340"

    ws.Cells(rowNum, "D").value = "ROBINSONS"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 23

    ws.Cells(rowNum, "B").value = 23

    ws.Cells(rowNum, "C").value = "000-123-826"

    ws.Cells(rowNum, "D").value = "GRAND UNION SUPERMARKET INC"

    ws.Cells(rowNum, "E").value = "FILINVEST, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 24

    ws.Cells(rowNum, "B").value = 24

    ws.Cells(rowNum, "C").value = "008-188-059"

    ws.Cells(rowNum, "D").value = "KNOXPORT INC"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG"

    ws.Cells(rowNum, "F").value = "Miscellaneous"

    rowNum = rowNum + 1

    

    ' 25

    ws.Cells(rowNum, "B").value = 25

    ws.Cells(rowNum, "C").value = "009-754-702"

    ws.Cells(rowNum, "D").value = "GOOD SMALL SHEEP INC"

    ws.Cells(rowNum, "E").value = "DAANG BAKAL, MANDALUYONG"

    ws.Cells(rowNum, "F").value = "Miscellaneous"

    rowNum = rowNum + 1

    

    ' 26

    ws.Cells(rowNum, "B").value = 26

    ws.Cells(rowNum, "C").value = "163-955-615"

    ws.Cells(rowNum, "D").value = "725 TAGTEAM FUEL STATION"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Fuel and Oil"

    rowNum = rowNum + 1

    

    ' 27

    ws.Cells(rowNum, "B").value = 27

    ws.Cells(rowNum, "C").value = "007-875-333"

    ws.Cells(rowNum, "D").value = "SOUTHERN ONGVILLE CORP"

    ws.Cells(rowNum, "E").value = "CANDELARIA, QUEZON"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 28

    ws.Cells(rowNum, "B").value = 28

    ws.Cells(rowNum, "C").value = "905-035-955"

    ws.Cells(rowNum, "D").value = "N MART GENERAL MERCHANDIZE"

    ws.Cells(rowNum, "E").value = "BAMBANG, NUEVA ECIJA"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 29

    ws.Cells(rowNum, "B").value = 29

    ws.Cells(rowNum, "C").value = "241-838-060"

    ws.Cells(rowNum, "D").value = "FACIAL CARE CENTRE"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 30

    ws.Cells(rowNum, "B").value = 30

    ws.Cells(rowNum, "C").value = "009-149--884"

    ws.Cells(rowNum, "D").value = "LANDMARK SUPERMARKET"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 31

    ws.Cells(rowNum, "B").value = 31

    ws.Cells(rowNum, "C").value = "240-940-697"

    ws.Cells(rowNum, "D").value = "INENGS"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 32

    ws.Cells(rowNum, "B").value = 32

    ws.Cells(rowNum, "C").value = "009-271-384"

    ws.Cells(rowNum, "D").value = "OTADER DENTAL AND MEDICAL"

    ws.Cells(rowNum, "E").value = "MALATE, MANILA"

    ws.Cells(rowNum, "F").value = "Clinic Materials and Supplies"

    rowNum = rowNum + 1

    

    ' 33

    ws.Cells(rowNum, "B").value = 33

    ws.Cells(rowNum, "C").value = "000-340-727"

    ws.Cells(rowNum, "D").value = "ALPHADENT CORPORATION"

    ws.Cells(rowNum, "E").value = "SAN ANTONIO, QUEZON"

    ws.Cells(rowNum, "F").value = "Clinic Materials and Supplies"

    rowNum = rowNum + 1

    

    ' 34

    ws.Cells(rowNum, "B").value = 34

    ws.Cells(rowNum, "C").value = "006-614-702"

    ws.Cells(rowNum, "D").value = "PROS APAC CORP"

    ws.Cells(rowNum, "E").value = "MALATE, MANILA"

    ws.Cells(rowNum, "F").value = "Clinic Materials and Supplies"

    rowNum = rowNum + 1

    

    ' 35

    ws.Cells(rowNum, "B").value = 35

    ws.Cells(rowNum, "C").value = "003-935-334"

    ws.Cells(rowNum, "D").value = "ST PATRICK MEDICAL SYSTEMS INC"

    ws.Cells(rowNum, "E").value = "SHAW BLVD., MANDALUYONG"

    ws.Cells(rowNum, "F").value = "Clinic Materials and Supplies"

    rowNum = rowNum + 1

    

    ' 36

    ws.Cells(rowNum, "B").value = 36

    ws.Cells(rowNum, "C").value = "216-037-865"

    ws.Cells(rowNum, "D").value = "DENTAL DOMAIN CORP"

    ws.Cells(rowNum, "E").value = "CULIAT, QUEZON CITY"

    ws.Cells(rowNum, "F").value = "Clinic Materials and Supplies"

    rowNum = rowNum + 1

    

    ' 37

    ws.Cells(rowNum, "B").value = 37

    ws.Cells(rowNum, "C").value = "226-527-915"

    ws.Cells(rowNum, "D").value = "METRO MARKET MARKET"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 38

    ws.Cells(rowNum, "B").value = 38

    ws.Cells(rowNum, "C").value = "010-176-614"

    ws.Cells(rowNum, "D").value = "GREEN DENTAL LAB"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG"

    ws.Cells(rowNum, "F").value = "Clinic Materials and Supplies"

    rowNum = rowNum + 1

    

    ' 39

    ws.Cells(rowNum, "B").value = 39

    ws.Cells(rowNum, "C").value = "440-275-225"

    ws.Cells(rowNum, "D").value = "AAHA GAS RETAILING"

    ws.Cells(rowNum, "E").value = "PAMPLONA UNO, LAS PIÑAS"

    ws.Cells(rowNum, "F").value = "Fuel and Oil"

    rowNum = rowNum + 1

    

    ' 40

    ws.Cells(rowNum, "B").value = 40

    ws.Cells(rowNum, "C").value = "010-119-810"

    ws.Cells(rowNum, "D").value = "RJD MERRY HARVEST CORP"

    ws.Cells(rowNum, "E").value = "SILANG, CAVITE"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 41

    ws.Cells(rowNum, "B").value = 41

    ws.Cells(rowNum, "C").value = "000-522-680"

    ws.Cells(rowNum, "D").value = "ALBA INTERNATIONAL INC"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Fuel and Oil"

    rowNum = rowNum + 1

    

    ' 42

    ws.Cells(rowNum, "B").value = 42

    ws.Cells(rowNum, "C").value = "007-551-518"

    ws.Cells(rowNum, "D").value = "NUTRITION FOR LIFE"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 43

    ws.Cells(rowNum, "B").value = 43

    ws.Cells(rowNum, "C").value = "008-664-864"

    ws.Cells(rowNum, "D").value = "GRANDWEN GAS CORP"

    ws.Cells(rowNum, "E").value = "BACLARAN, PASAY"

    ws.Cells(rowNum, "F").value = "Fuel and Oil"

    rowNum = rowNum + 1

    

    ' 44

    ws.Cells(rowNum, "B").value = 44

    ws.Cells(rowNum, "C").value = "009-195-748"

    ws.Cells(rowNum, "D").value = "BIG SKY NATION INC"

    ws.Cells(rowNum, "E").value = "183 PASAY"

    ws.Cells(rowNum, "F").value = "Fuel and Oil"

    rowNum = rowNum + 1

    

    ' 45

    ws.Cells(rowNum, "B").value = 45

    ws.Cells(rowNum, "C").value = "241-437-266"

    ws.Cells(rowNum, "D").value = "YELOW CAB PIZZA"

    ws.Cells(rowNum, "E").value = "STA. ROSA, LAGUNA"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 46

    ws.Cells(rowNum, "B").value = 46

    ws.Cells(rowNum, "C").value = "636-985956"

    ws.Cells(rowNum, "D").value = "WILSON MOA"

    ws.Cells(rowNum, "E").value = "MOA, PASAY"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 47

    ws.Cells(rowNum, "B").value = 47

    ws.Cells(rowNum, "C").value = "206-114-170"

    ws.Cells(rowNum, "D").value = "PETRON EXPRESS CENTER"

    ws.Cells(rowNum, "E").value = "BALAGTAS, BULACAN"

    ws.Cells(rowNum, "F").value = "Fuel and Oil"

    rowNum = rowNum + 1

    

    ' 48

    ws.Cells(rowNum, "B").value = 48

    ws.Cells(rowNum, "C").value = "005-148-299"

    ws.Cells(rowNum, "D").value = "ERNA TRADINGCORP"

    ws.Cells(rowNum, "E").value = "TALON 4, LAS PIÑAS"

    ws.Cells(rowNum, "F").value = "Repairs and Maintenance"

    rowNum = rowNum + 1

    

    ' 49

    ws.Cells(rowNum, "B").value = 49

    ws.Cells(rowNum, "C").value = "008-367-582"

    ws.Cells(rowNum, "D").value = "JAVELLANA ADVANCED AESTHETICS CLINIC"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Clinic Materials and Supplies"

    rowNum = rowNum + 1

    

    ' 50

    ws.Cells(rowNum, "B").value = 50

    ws.Cells(rowNum, "C").value = "226-784-068-000"

    ws.Cells(rowNum, "D").value = "MDI Group Holdings Inc"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG CITY"

    ws.Cells(rowNum, "F").value = "Rental"

    rowNum = rowNum + 1

    

    ' 51

    ws.Cells(rowNum, "B").value = 51

    ws.Cells(rowNum, "C").value = "000-405-340"

    ws.Cells(rowNum, "D").value = "ROBINSONS"

    ws.Cells(rowNum, "E").value = "TALON UNO, LAS PIÑAS"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 52

    ws.Cells(rowNum, "B").value = 52

    ws.Cells(rowNum, "C").value = "007-189-834"

    ws.Cells(rowNum, "D").value = "KENNY ROGERS"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 53

    ws.Cells(rowNum, "B").value = 53

    ws.Cells(rowNum, "C").value = "001-096-140"

    ws.Cells(rowNum, "D").value = "ALMON MARINA INC"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 54

    ws.Cells(rowNum, "B").value = 54

    ws.Cells(rowNum, "C").value = "236-414-781"

    ws.Cells(rowNum, "D").value = "YOKI FARM"

    ws.Cells(rowNum, "E").value = "PALOCPOC, CAVITE"

    ws.Cells(rowNum, "F").value = "Miscellaneous"

    rowNum = rowNum + 1

    

    ' 55

    ws.Cells(rowNum, "B").value = 55

    ws.Cells(rowNum, "C").value = "203-301-457"

    ws.Cells(rowNum, "D").value = "CINNABON"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 56

    ws.Cells(rowNum, "B").value = 56

    ws.Cells(rowNum, "C").value = "000-388-474"

    ws.Cells(rowNum, "D").value = "MERCURY"

    ws.Cells(rowNum, "E").value = "TALON 5, LAS PIÑAS"

    ws.Cells(rowNum, "F").value = "Clinic Materials and Supplies"

    rowNum = rowNum + 1

    

    ' 57

    ws.Cells(rowNum, "B").value = 57

    ws.Cells(rowNum, "C").value = "000-123-826"

    ws.Cells(rowNum, "D").value = "SOUTH SUPERMARKET"

    ws.Cells(rowNum, "E").value = "FILINVEST, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 58

    ws.Cells(rowNum, "B").value = 58

    ws.Cells(rowNum, "C").value = "241-498-356"

    ws.Cells(rowNum, "D").value = "HONEYBON"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 59

    ws.Cells(rowNum, "B").value = 59

    ws.Cells(rowNum, "C").value = "000-388-771"

    ws.Cells(rowNum, "D").value = "JOLLIBEE"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 60

    ws.Cells(rowNum, "B").value = 60

    ws.Cells(rowNum, "C").value = "610-884-579"

    ws.Cells(rowNum, "D").value = "T RIPLE MLD CORP"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Transportation and Travel"

    rowNum = rowNum + 1

    

    ' 61

    ws.Cells(rowNum, "B").value = 61

    ws.Cells(rowNum, "C").value = "607-237-567"

    ws.Cells(rowNum, "D").value = "RUSTIC PIZZA PASTA RESTAURANT INC"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 62

    ws.Cells(rowNum, "B").value = 62

    ws.Cells(rowNum, "C").value = "005-148-299"

    ws.Cells(rowNum, "D").value = "ERNA TRADING CORP"

    ws.Cells(rowNum, "E").value = "TALON 4, LAS PIÑAS"

    ws.Cells(rowNum, "F").value = "Clinic Materials and Supplies"

    rowNum = rowNum + 1

    

    ' 63

    ws.Cells(rowNum, "B").value = 63

    ws.Cells(rowNum, "C").value = "010-892-033"

    ws.Cells(rowNum, "D").value = "ELEVATE PILATES COMPANY LIMITED INC"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG"

    ws.Cells(rowNum, "F").value = "Clinic Materials and Supplies"

    rowNum = rowNum + 1

    

    ' 64

    ws.Cells(rowNum, "B").value = 64

    ws.Cells(rowNum, "C").value = "200-005-881"

    ws.Cells(rowNum, "D").value = "ARISTOCRAT BAKESHOP"

    ws.Cells(rowNum, "E").value = "ALMANZA UNO, LAS PIÑAS"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 65

    ws.Cells(rowNum, "B").value = 65

    ws.Cells(rowNum, "C").value = "000-060-696"

    ws.Cells(rowNum, "D").value = "RUSTANS"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Representation"

    rowNum = rowNum + 1

    

    ' 66

    ws.Cells(rowNum, "B").value = 66

    ws.Cells(rowNum, "C").value = "909-079-780"

    ws.Cells(rowNum, "D").value = "INGUSAN GASOLINE STATION"

    ws.Cells(rowNum, "E").value = "TALON UNO, LAS PIÑAS"

    ws.Cells(rowNum, "F").value = "Fuel and Oil"

    rowNum = rowNum + 1

    

    ' 67

    ws.Cells(rowNum, "B").value = 67

    ws.Cells(rowNum, "C").value = "000-299-299"

    ws.Cells(rowNum, "D").value = "ABACUS BOOK AND CARD CORP"

    ws.Cells(rowNum, "E").value = "POBLACION, MAKATI"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 68

    ws.Cells(rowNum, "B").value = 68

    ws.Cells(rowNum, "C").value = "003-058-789"

    ws.Cells(rowNum, "D").value = "SM PRIME HOLDINGS INC"

    ws.Cells(rowNum, "E").value = "ALMANZA UNO, LAS PIÑAS"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 69

    ws.Cells(rowNum, "B").value = 69

    ws.Cells(rowNum, "C").value = "183-378-131"

    ws.Cells(rowNum, "D").value = "MG PETRON CORDON SERVICE STATION"

    ws.Cells(rowNum, "E").value = "CORDON, ISABELA"

    ws.Cells(rowNum, "F").value = "Fuel and Oil"

    rowNum = rowNum + 1

    

    ' 70

    ws.Cells(rowNum, "B").value = 70

    ws.Cells(rowNum, "C").value = "630-167-414"

    ws.Cells(rowNum, "D").value = "NAIL GARDEN"

    ws.Cells(rowNum, "E").value = "76 PASAY"

    ws.Cells(rowNum, "F").value = "Miscellaneous"

    rowNum = rowNum + 1

    

    ' 71

    ws.Cells(rowNum, "B").value = 71

    ws.Cells(rowNum, "C").value = "230-577-870"

    ws.Cells(rowNum, "D").value = "LALONG FUEL STATION"

    ws.Cells(rowNum, "E").value = "SAN PABLO, LAGUNA"

    ws.Cells(rowNum, "F").value = "Fuel and Oil"

    rowNum = rowNum + 1

    

    ' 72

    ws.Cells(rowNum, "B").value = 72

    ws.Cells(rowNum, "C").value = "008-909-992"

    ws.Cells(rowNum, "D").value = "LANDERS"

    ws.Cells(rowNum, "E").value = "TAMBO, PARAÑAQUE"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 73

    ws.Cells(rowNum, "B").value = 73

    ws.Cells(rowNum, "C").value = "207-961-175"

    ws.Cells(rowNum, "D").value = "FESTIVAL MALL"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 74

    ws.Cells(rowNum, "B").value = 74

    ws.Cells(rowNum, "C").value = "010-262-979"

    ws.Cells(rowNum, "D").value = "LIN ET VIE PELLE MERCHANDISING"

    ws.Cells(rowNum, "E").value = "ALABANG, MUNTINLUPA"

    ws.Cells(rowNum, "F").value = "Miscellaneous"

    rowNum = rowNum + 1

    

    ' 75

    ws.Cells(rowNum, "B").value = 75

    ws.Cells(rowNum, "C").value = "005-695-758"

    ws.Cells(rowNum, "D").value = "UNICITY"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG"

    ws.Cells(rowNum, "F").value = "Miscellaneous"

    rowNum = rowNum + 1

    

    ' 76

    ws.Cells(rowNum, "B").value = 76

    ws.Cells(rowNum, "C").value = "010-176-614"

    ws.Cells(rowNum, "D").value = "GREEN DELTA LAB PHILIPPINES INC"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG"

    ws.Cells(rowNum, "F").value = "Clinic Materials and Supplies"

    rowNum = rowNum + 1

    

    ' 77

    ws.Cells(rowNum, "B").value = 77

    ws.Cells(rowNum, "C").value = "260-019-682"

    ws.Cells(rowNum, "D").value = "BURDA MEDICAL APPAREL"

    ws.Cells(rowNum, "E").value = "SANTA CRUZ, MANILA"

    ws.Cells(rowNum, "F").value = "Clinic Materials and Supplies"

    rowNum = rowNum + 1

    

    ' 78

    ws.Cells(rowNum, "B").value = 78

    ws.Cells(rowNum, "C").value = "278-353-148"

    ws.Cells(rowNum, "D").value = "D AND L MEDICAL APPAREL"

    ws.Cells(rowNum, "E").value = "SANTA CRUZ, MANILA"

    ws.Cells(rowNum, "F").value = "Clinic Materials and Supplies"

    rowNum = rowNum + 1

    

    ' 79

    ws.Cells(rowNum, "B").value = 79

    ws.Cells(rowNum, "C").value = "778-866-700"

    ws.Cells(rowNum, "D").value = "GREEN WAVE COMPUTER SYSTEM"

    ws.Cells(rowNum, "E").value = "SANTA CRUZ, MANILA"

    ws.Cells(rowNum, "F").value = "Repairs and Maintenance"

    rowNum = rowNum + 1

    

    ' 80

    ws.Cells(rowNum, "B").value = 80

    ws.Cells(rowNum, "C").value = "010-559-959"

    ws.Cells(rowNum, "D").value = "FRONTIER DENTAL PRODUCTS CORP"

    ws.Cells(rowNum, "E").value = "SAN ANTONIO, MAKATI"

    ws.Cells(rowNum, "F").value = "Clinic Materials and Supplies"

    rowNum = rowNum + 1

    

    ' 81

    ws.Cells(rowNum, "B").value = 81

    ws.Cells(rowNum, "C").value = "000-324-286"

    ws.Cells(rowNum, "D").value = "GOLDEN PEAK SALES CORP"

    ws.Cells(rowNum, "E").value = "SAN NICOLAS, MANILA"

    ws.Cells(rowNum, "F").value = "Miscellaneous"

    rowNum = rowNum + 1

    

    ' 82

    ws.Cells(rowNum, "B").value = 82

    ws.Cells(rowNum, "C").value = "225-527-914"

    ws.Cells(rowNum, "D").value = "METRO RETAIL STORES GROUP INC"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 83

    ws.Cells(rowNum, "B").value = 83

    ws.Cells(rowNum, "C").value = "208-544-458"

    ws.Cells(rowNum, "D").value = "MKYR TRANSPORT SERVICES"

    ws.Cells(rowNum, "E").value = "BATASAN HILLS, QUEZON CITY"

    ws.Cells(rowNum, "F").value = "Transportation and Travel"

    rowNum = rowNum + 1

    

    ' 84

    ws.Cells(rowNum, "B").value = 84

    ws.Cells(rowNum, "C").value = "006-614-702"

    ws.Cells(rowNum, "D").value = "PROS-APAC CORP"

    ws.Cells(rowNum, "E").value = "MALATE, MANILA"

    ws.Cells(rowNum, "F").value = "Clinic Materials and Supplies"

    rowNum = rowNum + 1

    

    ' 85

    ws.Cells(rowNum, "B").value = 85

    ws.Cells(rowNum, "C").value = "201-619-152"

    ws.Cells(rowNum, "D").value = "VILLACORTA DENTAL LABORATY INC"

    ws.Cells(rowNum, "E").value = "SANTA ROSA, LAGUNA"

    ws.Cells(rowNum, "F").value = "Clinic Materials and Supplies"

    rowNum = rowNum + 1

    

    ' ROW 86 REMOVED - DUPLICATE

    

    ' 87

    ws.Cells(rowNum, "B").value = 87

    ws.Cells(rowNum, "C").value = "120-568-073"

    ws.Cells(rowNum, "D").value = "7-ELEVEN"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG"

    ws.Cells(rowNum, "F").value = "Miscellaneous"

    rowNum = rowNum + 1

    

    ' 88

    ws.Cells(rowNum, "B").value = 88

    ws.Cells(rowNum, "C").value = "000-388-474"

    ws.Cells(rowNum, "D").value = "MERCURY DRUG"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 89

    ws.Cells(rowNum, "B").value = 89

    ws.Cells(rowNum, "C").value = "233-251-708"

    ws.Cells(rowNum, "D").value = "FEDERAL BRENT RETAIL INC"

    ws.Cells(rowNum, "E").value = "076 PASAY"

    ws.Cells(rowNum, "F").value = "Supplies"

    rowNum = rowNum + 1

    

    ' 90

    ws.Cells(rowNum, "B").value = 90

    ws.Cells(rowNum, "C").value = "214-706-591"

    ws.Cells(rowNum, "D").value = "WATSONS"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG"

    ws.Cells(rowNum, "F").value = "Miscellaneous"

    rowNum = rowNum + 1

    

    ' 91

    ws.Cells(rowNum, "B").value = 91

    ws.Cells(rowNum, "C").value = "226-784-068"

    ws.Cells(rowNum, "D").value = "MDI Group Holdings Inc 2"

    ws.Cells(rowNum, "E").value = "BGC, TAGUIG"

    ws.Cells(rowNum, "F").value = "Association Dues"

    rowNum = rowNum + 1

    

    ' 92

    ws.Cells(rowNum, "B").value = 92

    ws.Cells(rowNum, "C").value = "TIN NOT FOUND"

    ws.Cells(rowNum, "D").value = "Rosalina Ignacio"

    ws.Cells(rowNum, "E").value = "—"

    ws.Cells(rowNum, "F").value = "Clinician's Fee-Corporators"

    rowNum = rowNum + 1

    

    ' 93

    ws.Cells(rowNum, "B").value = 93

    ws.Cells(rowNum, "C").value = "TIN NOT FOUND"

    ws.Cells(rowNum, "D").value = "Daisy Cornejo"

    ws.Cells(rowNum, "E").value = "—"

    ws.Cells(rowNum, "F").value = "Clinician's Fee-Corporators"

    rowNum = rowNum + 1

    

    ' 94

    ws.Cells(rowNum, "B").value = 94

    ws.Cells(rowNum, "C").value = "TIN NOT FOUND"

    ws.Cells(rowNum, "D").value = "Janina Tayag"

    ws.Cells(rowNum, "E").value = "—"

    ws.Cells(rowNum, "F").value = "Clinician's Fee-Corporators"

    rowNum = rowNum + 1

    

    ' 95

    ws.Cells(rowNum, "B").value = 95

    ws.Cells(rowNum, "C").value = "TIN NOT FOUND"

    ws.Cells(rowNum, "D").value = "Jasper Tago"

    ws.Cells(rowNum, "E").value = "—"

    ws.Cells(rowNum, "F").value = "Clinician's Fee-Corporators"

    rowNum = rowNum + 1

    

    ' 96

    ws.Cells(rowNum, "B").value = 96

    ws.Cells(rowNum, "C").value = "TIN NOT FOUND"

    ws.Cells(rowNum, "D").value = "Ana Ramona Locsin"

    ws.Cells(rowNum, "E").value = "—"

    ws.Cells(rowNum, "F").value = "Clinician's Fee-Visiting"

    rowNum = rowNum + 1

    

    ' 97

    ws.Cells(rowNum, "B").value = 97

    ws.Cells(rowNum, "C").value = "TIN NOT FOUND"

    ws.Cells(rowNum, "D").value = "Virgilio Malenab Jr"

    ws.Cells(rowNum, "E").value = "—"

    ws.Cells(rowNum, "F").value = "Clinician's Fee-Visiting"

    rowNum = rowNum + 1

    

    ' 98

    ws.Cells(rowNum, "B").value = 98

    ws.Cells(rowNum, "C").value = "TIN NOT FOUND"

    ws.Cells(rowNum, "D").value = "SBM Consulting Inc"

    ws.Cells(rowNum, "E").value = "—"

    ws.Cells(rowNum, "F").value = "Professional Fee"

    rowNum = rowNum + 1

    

    ' 99

    ws.Cells(rowNum, "B").value = 99

    ws.Cells(rowNum, "C").value = "TIN NOT FOUND"

    ws.Cells(rowNum, "D").value = "Christian Anthony Ermita"

    ws.Cells(rowNum, "E").value = "—"

    ws.Cells(rowNum, "F").value = "Clinician's Fee-Visiting"

    rowNum = rowNum + 1

    

    ' 100

    ws.Cells(rowNum, "B").value = 100

    ws.Cells(rowNum, "C").value = "TIN NOT FOUND"

    ws.Cells(rowNum, "D").value = "Vergilyn Montales"

    ws.Cells(rowNum, "E").value = "—"

    ws.Cells(rowNum, "F").value = "Clinician's Fee-Visiting"

    rowNum = rowNum + 1

    

    ' ---- VALIDATE FOR DUPLICATES ----

    ' Check: Same Name + Same Address = NOT ALLOWED

    lastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).row

    Dim duplicateFound As Boolean

    duplicateFound = False

    

    For i = 2 To lastRow

        Dim name As String

        Dim address As String

        Dim j As Long

        

        name = ws.Cells(i, "D").value

        address = ws.Cells(i, "E").value

        

        If name <> "" And address <> "" And address <> "—" Then

            For j = i + 1 To lastRow

                Dim name2 As String

                Dim address2 As String

                

                name2 = ws.Cells(j, "D").value

                address2 = ws.Cells(j, "E").value

                

                If name = name2 And address = address2 Then

                    MsgBox "DUPLICATE FOUND:" & vbCrLf & _

                           "Row " & i & ": " & name & " | " & address & vbCrLf & _

                           "Row " & j & ": " & name2 & " | " & address2 & vbCrLf & vbCrLf & _

                           "Same Name + Address is NOT allowed!", vbCritical

                    duplicateFound = True

                    Exit Sub

                End If

            Next j

        End If

    Next i

    

    If Not duplicateFound Then

        MsgBox "Import complete! " & (rowNum - 2) & " suppliers imported." & vbCrLf & _

               "All suppliers have unique Name + Address combinations.", vbInformation

    End If

End Sub









