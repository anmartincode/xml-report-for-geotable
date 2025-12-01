============================================================================================================================================================================================================================
COMPLETE GEOTABLE FORMAT ANALYSIS
Green Line High Speed Test Track Project (GLTT) - 100% Submission

LABEL CONVENTION UPDATE (2025 IMPLEMENTATION)
-------------------------------------------
Runtime output now uses InRoads-style labels:
  POB (Point of Beginning) replaces legacy POT at initial tangent start.
  PC / PT for non-spiral circular curve start/end (was SC / CS in some legacy examples).
  PVC / PVI / PVT for vertical curve points.
Spiral transition labels (TS, SC, CS, ST) are preserved only when true spiral entities are present.
Documentation rows below retain original source spreadsheet labels for structural analysis reference.
============================================================================================================================================================================================================================


############################################################################################################################################################################################################################
SHEET 1: InRails Reader
############################################################################################################################################################################################################################
Dimensions: A1:M13
Max Row: 13, Max Column: 13

MERGED CELLS:
  F4:M4
  F9:M9
  F12:M12
  F11:M11
  F10:M10
  F6:M6
  F5:M5

DETAILED ROW STRUCTURE:
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Row 1:
  A=v1.5

Row 3:
  F=Procedure:

Row 4:
  F=1. Click the appropriate "InRails Reader" button to bring in an exported InRails spreadsheet on a new tab reformatted fo

Row 5:
  F=2. After the program has run, manually edit the curve number label, turnout number label and turnout size/RH&LH labels. 

Row 6:
  F=3. Click the "Import curve data" button that comes in on the new tab to import speed and superelevation data from horizo

Row 8:
  F=Notes:

Row 9:
  F=* Turnout labels/rectangles will only be inserted if "PI"s have been manually changed to "PITO" and "PS" before running 

Row 10:
  F=* If the InRails data file that you are importing has more than one tab, the only tab that will be imported is the first

Row 11:
  F=* If there are speed and superelevation values that don't fill in after running "Import curve data", check that the curv

Row 12:
  F=* If you are having a problem with the program closing when you try to import a file, try closing all other Excel files 


############################################################################################################################################################################################################################
SHEET 2: PROP WB CH
############################################################################################################################################################################################################################
Dimensions: A1:U12
Max Row: 12, Max Column: 21

MERGED CELLS:
  H5:K6
  F2:G2
  H4:K4
  C2:C3
  A1:K1
  A2:A3
  H7:K7
  B2:B3
  D2:D3
  H2:K3
  E2:E3
  B5:B6

DETAILED ROW STRUCTURE:
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Row 1:
  A=TRACK GEOMETRY DATA - WESTBOUND CHESTNUT HILL

Row 2:
  A=ELEMENT
  B=CURVE No.
  C=POINT
  D=STATION
  E=BEARING
  F=COORDINATES
  H=DATA

Row 3:
  F=Northing
  G=Easting

Row 4:
  A=TANGENT
  C=POL
  D=227+24.29
  E=N 75°32'18.69" W
  F=2944361.4598
  G=745856.9556
  H=L = 197.47'

Row 5:
  A=TANGENT
  B=17W
  C=PITO
  D=229+21.76
  E=N 75°32'18.69" W
  F=2944410.7739
  G=745665.7419
  H=NO. 10 RH CROSSOVER

Row 6:
  A=TANGENT
  C=PS
  D=229+53.18
  E=N 75°32'18.69" W
  F=2944418.6195
  G=745635.3206

Row 7:
  A=TANGENT
  C=POL
  D=236+98.99
  E=N 75°32'18.69" W
  F=2944604.8704
  G=744913.1389
  H=L = 745.81'


############################################################################################################################################################################################################################
SHEET 3: PROP WB NC
############################################################################################################################################################################################################################
Dimensions: A1:U12
Max Row: 12, Max Column: 21

MERGED CELLS:
  H4:K4
  F2:G2
  H5:K6
  A8:A9
  E8:E9
  B10:B12
  C2:C3
  E10:E12
  A1:K1
  B8:B9
  A2:A3
  H7:K7
  B2:B3
  D2:D3
  H2:K3
  E2:E3
  B5:B6

DETAILED ROW STRUCTURE:
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Row 1:
  A=TRACK GEOMETRY DATA - WESTBOUND NEWTON CENTRE

Row 2:
  A=ELEMENT
  B=CURVE No.
  C=POINT
  D=STATION
  E=BEARING
  F=COORDINATES
  H=DATA

Row 3:
  F=Northing
  G=Easting

Row 4:
  A=TANGENT
  C=POL
  D=268+68.17
  E=N 75°32'18.69" W
  F=2945396.3068
  G=741844.3704
  H=L = 32.97'

Row 5:
  A=TANGENT
  B=18E
  C=PS
  D=269+01.14
  E=N 75°32'18.69" W
  F=2945404.5399
  G=741812.447
  H=NO. 10 LH CROSSOVER

Row 6:
  A=TANGENT
  C=PITO
  D=269+32.56
  E=N 75°32'18.69" W
  F=2945412.3856
  G=741782.0258

Row 7:
  A=TANGENT
  E=N 75°32'18.69" W
  H=L = 264.77'

Row 8:
  A=SPIRAL
  C=TS
  D=271+97.33
  F=2945478.5076
  G=741525.6399
  H=θs = 2°14'01.95"
  I=Ls= 210.00'
  J=LT= 140.01'
  K=STs= 70.01'

Row 9:
  H=Xs= 209.97'
  I=Ys= 2.73'
  J=P= 0.68'
  K=K= 104.99'

Row 10:
  B=W-C-1
  C=SC
  D=274+07.33
  F=2945528.3002
  G=741321.643
  H=Δc = 0°00'22.27"
  I=Da= 2°07'39.00"
  J=R= 2693.10'
  K=Lc= 0.29'

Row 11:
  A=CURVE
  C=PI
  F=2945528.331
  G=741321.5009
  H=V= 40 MPH
  I=Ea= 2.00"
  J=Ee= 2.38"
  K=Eu= 0.38"

Row 12:
  C=POC
  D=274+07.62
  F=2945528.3618
  G=741321.3588
  H=Tc= 0.15'
  I=Ec= 0.00'
  J=CC:N 2942896.2932
  K=E 740751.2557


############################################################################################################################################################################################################################
SHEET 4: PROP EB CH
############################################################################################################################################################################################################################
Dimensions: A1:U12
Max Row: 12, Max Column: 21

MERGED CELLS:
  H10:K11
  F2:G2
  H9:K9
  B4:B6
  C2:C3
  H12:K12
  E4:E6
  A7:A8
  B7:B8
  E7:E8
  A2:A3
  B2:B3
  B10:B11
  D2:D3
  E2:E3
  H2:K3
  A1:K1

DETAILED ROW STRUCTURE:
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Row 1:
  A=TRACK GEOMETRY DATA - EASTBOUND (TT) CHESTNUT HILL
  M=            <--- use this button AFTER filling in curve labels.

Row 2:
  A=ELEMENT
  B=CURVE No.
  C=POINT
  D=STATION
  E=BEARING
  F=COORDINATES
  H=DATA

Row 3:
  F=Northing
  G=Easting

Row 4:
  B=E-C-1
  C=POC
  D=222+73.46
  F=2944253.4698
  G=746294.2396
  H=Δc = 2°17'31.84"
  I=Da= 3°06'42.58"
  J=R= 1841.23'
  K=Lc= 73.66'

Row 5:
  A=CURVE
  C=PI
  F=2944258.0307
  G=746257.6881
  H=V= 40 MPH
  I=Ea= 2.00"
  J=Ee= 3.49"
  K=Eu= 1.49"

Row 6:
  C=CS
  D=223+47.12
  F=2944264.0498
  G=746221.3482
  H=Tc= 36.83'
  I=Ec= 0.37'
  J=CC:N 2946080.5272
  K=E 746522.2190

Row 7:
  A=SPIRAL
  H=θs = 5°03'24.19"
  I=Ls= 325.00'
  J=LT= 216.76'
  K=STs= 108.41'

Row 8:
  C=ST
  D=226+72.12
  F=2944335.8955
  G=745904.5043
  H=Xs= 324.75'
  I=Ys= 9.56'
  J=P= 2.39'
  K=K= 162.46'

Row 9:
  A=TANGENT
  E=N 75°32'18.69" W
  H=L = 90.00'

Row 10:
  A=TANGENT
  B=17E
  C=PS
  D=227+62.12
  E=N 75°32'18.69" W
  F=2944358.3711
  G=745817.3558
  H=NO. 10 RH CROSSOVER

Row 11:
  A=TANGENT
  C=PITO
  D=227+93.53
  E=N 75°32'18.69" W
  F=2944366.2168
  G=745786.9346

Row 12:
  A=TANGENT
  C=POL
  D=231+58.72
  E=N 75°32'18.69" W
  F=2944457.4154
  G=745433.3151
  H=L = 365.19'


############################################################################################################################################################################################################################
SHEET 5: PROP EB NC
############################################################################################################################################################################################################################
Dimensions: A1:U12
Max Row: 12, Max Column: 21

MERGED CELLS:
  H4:K4
  F2:G2
  H5:K6
  A8:A9
  E8:E9
  B10:B12
  C2:C3
  E10:E12
  A1:K1
  B8:B9
  A2:A3
  H7:K7
  B2:B3
  D2:D3
  H2:K3
  E2:E3
  B5:B6

DETAILED ROW STRUCTURE:
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Row 1:
  A=TRACK GEOMETRY DATA - EASTBOUND (TT) NEWTON CENTRE
  M=            <--- use this button AFTER filling in curve labels.

Row 2:
  A=ELEMENT
  B=CURVE No.
  C=POINT
  D=STATION
  E=BEARING
  F=COORDINATES
  H=DATA

Row 3:
  F=Northing
  G=Easting

Row 4:
  A=TANGENT
  C=POL
  D=268+68.43
  E=N 75°32'18.69" W
  F=2945383.8358
  G=741841.1506
  H=L = 192.86'

Row 5:
  A=TANGENT
  B=18W
  C=PITO
  D=270+61.29
  E=N 75°32'18.69" W
  F=2945431.9988
  G=741654.4001
  H=NO. 10 LH CROSSOVER

Row 6:
  A=TANGENT
  C=PS
  D=270+92.71
  E=N 75°32'18.69" W
  F=2945439.8445
  G=741623.9788

Row 7:
  A=TANGENT
  E=N 75°32'18.69" W
  H=L = 90.00'

Row 8:
  A=SPIRAL
  C=TS
  D=271+82.71
  F=2945462.3201
  G=741536.8304
  H=θs = 1°56'02.01"
  I=Ls= 200.00'
  J=LT= 133.34'
  K=STs= 66.67'

Row 9:
  H=Xs= 199.98'
  I=Ys= 2.25'
  J=P= 0.56'
  K=K= 100.00'

Row 10:
  B=E-C-2
  C=SC
  D=273+82.71
  F=2945510.0815
  G=741342.6274
  H=Δc = 0°05'40.67"
  I=Da= 1°56'02.01"
  J=R= 2962.72'
  K=Lc= 4.89'

Row 11:
  A=CURVE
  C=PI
  F=2945510.6122
  G=741340.239
  H=V= 40 MPH
  I=Ea= 2.00"
  J=Ee= 2.17"
  K=Eu= 0.17"

Row 12:
  C=POC
  D=273+87.60
  F=2945511.1389
  G=741337.8498
  H=Tc= 2.45'
  I=Ec= 0.00'
  J=CC:N 2942617.8991
  K=E 740699.9851