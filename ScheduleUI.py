from gooey import Gooey, GooeyParser
import os
import mkl
import json
"""
Program: ScheduleUI.py
Purpose: Handles front end of VolunteerOrganizer application. With a Gooey
based UI asking for volunteer positions, times, shift types, etc. For venue
grid & shift list generation
"""
@Gooey(program_name="Venue Grid Generator")
def parse_args():
    """ Use GooeyParser to build up the arguments we will use in our script
    Save the arguments in a default json file so that we can retrieve them
    every time we run the script.
    """
    stored_args = {}
    # get the script name without the extension & use it to build up the json filename
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    args_file = "{}-args.json".format(script_name)
    # Read in the prior arguments as a dictionary
    if os.path.isfile(args_file):
        with open(args_file) as data_file:
            stored_args = json.load(data_file)
    parser = GooeyParser(description='Produce a Venue Schedule')
    #Creating Dimensions for schedule grid

    #Target Excel file, either can select a pre-existing .xlsm file or generate a new one
    parser.add_argument('--FileName',
                        metavar='File Name',
                        help = 'Select the target Microsoft Excel file (MUST be macro-enabled, as a .xlsm file).',
                        widget ='FileChooser' )
    parser.add_argument("-o",
                        "--OutputFile",
                        required=False,
                        help="Select a path for new Microsoft Excel File.",
                        widget="FileSaver",
                        metavar = "Output File",
                        gooey_options=dict(wildcard="Excel Macro-Enabled Workbook (.xlsm)|*.xlsm"))
    #Venue Name
    parser.add_argument('VenueName',
                        action='store',
                        default=stored_args.get('Venue_Name'),
                        metavar = "Venue Name",
                        help='Specify the name of the venue. No longer than 31 characters in length.')
    #Shift Time
    parser.add_argument('ShiftTime',
                        action='store',
                        default=stored_args.get('Shift_Time'),
                        metavar = "Shift Time",
                        help='Specify the approximate shift length.')
    #Universal Shift Time Check
    parser.add_argument('--UniversalShiftTime',
                        default=stored_args.get('Universal_Shift_Time'),
                        help='Will this venue have a single shift time?',
                        widget = 'CheckBox',
                        metavar = "Universal Shift Time",
                        action='store_true')
    #Venue Positions, selected based on # of vols (if 0, then role is not needed)
    parser.add_argument('--BGVols',
                action='store',
                widget='IntegerField',
                default=0,
                metavar = "Beer Garden Volunteers",
                help="Specify the number of beer garden volunteers that will be present.")
    parser.add_argument('--FOHVols',
            action='store',
            widget='IntegerField',
            default=0,
            metavar="Front of House Volunteers",
            help="Specify the number of front of house volunteers that will be present.")
    parser.add_argument('--GTVols',
            action='store',
            widget='IntegerField',
            default=0,
            metavar="Green Team Volunteers",
            help="Specify the number of green team volunteers that will be present.")
    parser.add_argument('--HospVols',
            action='store',
            widget='IntegerField',
            default=0,
            metavar="Hospitality Volunteers",
            help="Specify the number of hospitality volunteers that will be present.")
    parser.add_argument('--MerchVols',
            action='store',
            widget='IntegerField',
            default=0,
            metavar="Merchandise Volunteers",
            help="Specify the number of merchandise volunteers that will be present.")
    parser.add_argument('--SiteVols',
            action='store',
            widget='IntegerField',
            default=0,
            metavar="Site Crew Volunteers",
            help="Specify the number of site crew volunteers that will be present.")
    parser.add_argument('--SecVols',
            action='store',
            widget='IntegerField',
            default=0,
            metavar="Security Volunteers",
            help="Specify the number of security volunteers that will be present.")
    parser.add_argument('--FirstAidVols',
            action='store',
            widget='IntegerField',
            default=0,
            metavar = "First Aid Volunteers",
            help="Specify the number of first aid volunteers that will be present.")
    parser.add_argument('--SetupTeardownVols',
            action='store',
            widget='IntegerField',
            default=0,
            metavar = "Setup/Teardown Volunteers",
            help="Specify the number of setup/teardown volunteers that will be present.")
    parser.add_argument('--OfficeVols',
            action='store',
            widget='IntegerField',
            default=0,
            metavar="Office Volunteers",
            help="Specify the number of office volunteers that will be present.")
    parser.add_argument('--SurveyVols',
            action='store',
            widget='IntegerField',
            default=0,
            metavar="Survey Volunteers",
            help="Specify the number of survey volunteers that will be present.")                           
    #Type of supervisors at venue
    parser.add_argument('--SupervisorPositions',
                        choices=['Beer Garden','Front of House','Green Team','Hospitality','Merchandise','Staging','Security'],
                        action='store',
                        default=stored_args.get('Supervisor_Positons'),
                        widget='Listbox',
                        nargs="+",
                        metavar="Supervisor Positons",
                        help='Specify the supervisor work positons for this venue.')
    #Split shift for positions
    parser.add_argument('--SplitShifts',
                        choices= ['Beer Garden','Front of House','Green Team','Hospitality','Merchandise','Site Crew','Security', 'First Aid','Setup/Teardown', 'Office','Survey'],
                        action='store',
                        default=stored_args.get('Split_Shifts'),
                        widget='Listbox',
                        nargs="+",
                        metavar="Split Shifts",
                        help='Specify which work positions will have split shifts.')
    #Number of shows taking place at venue
    parser.add_argument('--NumberofDays',
                        action='store',
                        default=1,
                        widget='IntegerField',
                        metavar="Number of Days",
                        help="Specify the number of days/shows that will occur here.")
    #Return parser function to main                    
    return parser.parse_args()
