

moveCells_OrgMaintenance = {'B24': 'AD', 'C24': 'AE', 'B25': 'AF', 'C25': 'AG', 'B26': 'AH', 'C26': 'AI', 'B27': 'AJ',
                            'C27': 'AK', 'B28': 'AL', 'C28': 'AM', 'B30': 'AN',
                            'C30': 'AO', 'B31': 'AP', 'C31': 'AQ', 'B32': 'AR', 'C32': 'AS', 'B33': 'AT', 'C33': 'AU',
                            'B35': 'AV', 'C35': 'AW', 'B37': 'AX'
                            }

# Programming columns
# Program Name, Event Description, Date, Attendance, Location, Admission, Room Rent and Equip., Advertising, Food, Supplies/Decorations, Duplications, Contracts(not included yet), Other,

moveCells_Programming = {'A2': ['BX', 'FS', 'JN'], 'A3': ['BY', 'FT', 'JO'], 'B3': ['BZ', 'FU', 'JP'],
                         'C3': ['CA', 'FV', 'JQ'],
                         'D3': ['CB', 'FW', 'JR'], 'E3': ['CC', 'FX', 'JS'], 'B5': ['CE', 'FZ', 'JU'],
                         'C5': ['CF', 'GA', 'JV'], 'B6': ['CG', 'GB', 'JW'],
                         'C6': ['CH', 'GC', 'JX'], 'B7': ['CI', 'GD', 'JY'], 'C7': ['CJ', 'GE', 'JZ'],
                         'B8': ['CK', 'GF', 'KA'], 'C8': ['CL', 'GG', 'KB'],
                         'B9': ['CM', 'GI', 'KE'], 'C9': ['CU', 'GQ', 'KM'], 'B10': ['CN', 'GJ', 'KF'],
                         'B11': ['CO', 'GK', 'KG'], 'B12': ['CP', 'GL', 'KH'],
                         'B13': ['CQ', 'GM', 'KI'], 'B14': ['CR', 'GN', 'KJ'], 'B15': ['CS', 'GO', 'KK'],
                         'B17': ['CV', 'GR', 'KN'], 'C17': ['CW', 'GS', 'KO'],
                         'B20': ['CX', 'GT', 'KP']
                         }


moveCells_Programming2 = {'G2': ['BX', 'FS', 'JN'], 'G3': ['BY', 'FT', 'JO'], 'H3': ['BZ', 'FU', 'JP'],
                          'I3': ['CA', 'FV', 'JQ'], 'J3': ['CB', 'FW', 'JR'], 'K3': ['CC', 'FX', 'JS'],
                          'H5': ['CE', 'FZ', 'JU'], 'I5': ['CF', 'GA', 'JV'], 'H6': ['CG', 'GB', 'JW'],
                          'I6': ['CH', 'GC', 'JX'], 'H7': ['CI', 'GD', 'JY'], 'I7': ['CJ', 'GE', 'JZ'],
                          'H8': ['CK', 'GF', 'KA'], 'I8': ['CL', 'GG', 'KB'], 'H9': ['CM', 'GI', 'KE'],
                          'I9': ['CU', 'GQ', 'KM'], 'H10': ['CN', 'GJ', 'KF'], 'H11': ['CO', 'GK', 'KG'],
                          'H12': ['CP', 'GL', 'KH'], 'H13': ['CQ', 'GM', 'KI'], 'H14': ['CR', 'GN', 'KJ'],
                          'H15': ['CS', 'GO', 'KK'], 'H17': ['CV', 'GR', 'KN'], 'I17': ['CW', 'GS', 'KO'],
                          'H20': ['CX', 'GT', 'KP']
                          }

# Series Programming Column

moveCells_seriesProgramming = {'H21': ['CZ', 'GW', 'KS'],  # Name
                               'H22': ['DA', 'GX', 'KT'],  # Description
                               'L21': ['DB', 'GV', 'KR'],  # Installments
                               'L22': ['DC', 'GY', 'KU'],  # Date
                               'K22': ['DD', 'GZ', 'KV'],  # Attendance
                               'J21': ['DE', 'HA', 'KW'],  # Location
                               'K21': ['DF', 'HB', 'KX'],  # admissions fee

                               'H24': ['DH', 'HD', 'KZ'],  # RRE
                               'I24': ['DI', 'HE', 'LA'],
                               'H25': ['DJ', 'HF', 'LB'],  # adv
                               'I25': ['DK', 'HG', 'LC'],
                               'H26': ['DL', 'HH', 'LD'],  # food
                               'I26': ['DM', 'HI', 'LE'],
                               'H27': ['DN', 'HJ', 'LF'],  # supply/deco
                               'I27': ['DO', 'HK', 'LG'],
                               'H28': ['DP', 'HL', 'LH'],
                               'G29': [['DQ', 'DR', 'DS', 'DT', 'DU', 'DV'], ['HM', 'HN', 'HO', 'HP', 'HQ', 'HR'],
                                       ['LI', 'LJ', 'LK', 'LL', 'LM', 'LN']],
                               # The concatenated columns, shows all the types of contracts requested
                               'I29': ['DX', 'HT', 'LP'],  # Description of Contract

                               'H32': ['DY', 'HU', 'LQ'],  # Other
                               'I32': ['DZ', 'HV', 'LR'],
                               }

# Trip Competition/Conference

moveCells_tripsCC1 = {'B50': ['EC', 'HY', 'LU'],  # NAME
                      'C52': ['ED', 'HZ', 'LV'],  # SERIES
                      'D52': ['EE', 'IA', 'LW'],  # LOCATION
                      'B52': ['EF', 'IB', 'LX'],  # DESCRIPTION
                      'E52': ['EG', 'IC', 'LY'],  # ATTENDANCE
                      'F52': ['EH', 'ID', 'LZ'],  # DATE

                      'B54': ['EJ', 'IF', 'MB'],  # TRANSPORT
                      'C54': ['EK', 'IG', 'MC'],
                      'B55': ['EL', 'IH', 'MD'],  # PARKING
                      'C55': ['EM', 'II', 'ME'],
                      'B56': ['EN', 'IJ', 'MF'],  # FOOD
                      'C56': ['EO', 'IK', 'MG'],
                      'B57': ['EP', 'IL', 'MH'],  # LODGING
                      'C57': ['EQ', 'IM', 'MI'],
                      'B58': ['ER', 'IN', 'MJ'],  # REGISTRATION
                      'C58': ['ES', 'IO', 'MK'],
                      'B59': ['ET', 'IP', 'ML'],  # OTHER
                      'C59': ['EU', 'IQ', 'MM'],
                      }


moveCells_TripsCC2 = {'H47': ['EC', 'HY', 'LU'],  # NAME
                      'I49': ['ED', 'HZ', 'LV'],  # SERIES
                      'J49': ['EE', 'IA', 'LW'],  # LOCATION
                      'H49': ['EF', 'IB', 'LX'],  # DESCRIPTION
                      'K49': ['EG', 'IC', 'LY'],  # ATTENDANCE
                      'L49': ['EH', 'ID', 'LZ'],  # DATE

                      'H51': ['EJ', 'IF', 'MB'],  # TRANSPORT
                      'I51': ['EK', 'IG', 'MC'],
                      'H52': ['EL', 'IH', 'MD'],  # PARKING
                      'I52': ['EM', 'II', 'ME'],
                      'H53': ['EN', 'IJ', 'MF'],  # FOOD
                      'I53': ['EO', 'IK', 'MG'],
                      'H54': ['EP', 'IL', 'MH'],  # LODGING
                      'I54': ['EQ', 'IM', 'MI'],
                      'H55': ['ER', 'IN', 'MJ'],  # REGISTRATION
                      'I55': ['ES', 'IO', 'MK'],
                      'H56': ['ET', 'IP', 'ML'],  # OTHER
                      'I56': ['EU', 'IQ', 'MM'],
                      }

# Other Trip
moveCells_otherTrip = {'H36': ['EX', 'IS', 'MP'],  # Name
                       'H35': ['EY', 'IT', 'MQ'],  # Series or nah
                       'J36': ['EZ', 'IU', 'MR'],  # Location
                       'H37': ['FA', 'IV', 'MS'],  # desc
                       'K37': ['FB', 'IW', 'MT'],  # atten
                       'L37': ['FC', 'IX', 'MU'],  # date

                       'H39': ['FE', 'JB', 'MY'],  # transp
                       'I39': ['FF', 'JC', 'MZ'],
                       'H40': ['FG', 'IZ', 'MW'],  # adv
                       'I40': ['FH', 'JA', 'MX'],
                       'H41': ['FI', 'JD', 'NA'],  # admin
                       'I41': ['FJ', 'JE', 'NB'],
                       'H42': ['FK', 'JF', 'NC'],  # food
                       'I42': ['FL', 'JG', 'ND'],
                       'H43': ['FM', 'JH', 'NE'],  # lodging
                       'I43': ['FN', 'JI', 'NF'],
                       'H44': ['FO', 'JJ', 'NG'],  # other
                       'I44': ['FP', 'JK', 'NH'],
                       }