from gooey import Gooey, GooeyParser
import os
import mkl
import json

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
    #File Name
    parser.add_argument('File_Name',
                        help = 'Select the target Microsoft Excel file (MUST be macro-enabled, as a .xlsm file).',
                        widget ='FileChooser' )
    #Venue Name
    parser.add_argument('Venue_Name',
                        action='store',
                        default=stored_args.get('Venue_Name'),
                        help='Specify the name of the venue, please use undercases for spaces')
    #Shift Time
    parser.add_argument('Shift_Time',
                        action='store',
                        default=stored_args.get('Shift_Time'),
                        help='Specify the approximate shift length')
    #Universal Shift Time Check
    parser.add_argument('--Universal_Shift_Time',
                        default=stored_args.get('Universal_Shift_Time'),
                        help='Will this venue have a single shift time?',
                        widget = 'CheckBox',
                        action='store_true')
    #Venue Positions, selected based on # of vols (if 0, then role is not needed)
    parser.add_argument('--BG_Vols',
                action='store',
                widget='IntegerField',
                default=0,
                help="Specify the number of BG volunteers that will be present")
    parser.add_argument('--FOH_Vols',
            action='store',
            widget='IntegerField',
            default=0,
            help="Specify the number of FOH volunteers that will be present")
    parser.add_argument('--GT_Vols',
            action='store',
            widget='IntegerField',
            default=0,
            help="Specify the number of Green Team volunteers that will be present")
    parser.add_argument('--Hosp_Vols',
            action='store',
            widget='IntegerField',
            default=0,
            help="Specify the number of hospitality volunteers that will be present")
    parser.add_argument('--Merch_Vols',
            action='store',
            widget='IntegerField',
            default=0,
            help="Specify the number of merchandise volunteers that will be present")
    parser.add_argument('--Stage_Vols',
            action='store',
            widget='IntegerField',
            default=0,
            help="Specify the number of staging volunteers that will be present")
    parser.add_argument('--Sec_Vols',
            action='store',
            widget='IntegerField',
            default=0,
            help="Specify the number of security volunteers that will be present")      
    #Type of supervisors at venue
    parser.add_argument('Supervisor_Positions',
                        choices=['Beer Garden','Front of House','Green Team','Hospitality','Merchandise','Staging','Security'],
                        action='store',
                        default=stored_args.get('Supervisor_Positons'),
                        widget='Listbox',
                        nargs="+",
                        metavar="Supervisor Positons",
                        help='Specify the work positons for this venue')
    #Split shift for positions
    parser.add_argument('--Split_Shifts',
                        choices=['Beer Garden','Front of House','Green Team','Hospitality','Merchandise','Staging','Security'],
                        action='store',
                        default=stored_args.get('Split_Shifts'),
                        widget='Listbox',
                        nargs="+",
                        metavar="Split_Shifts",
                        help='Specify which work positions will have split shifts')
    #Number of shows taking place at venue
    parser.add_argument('--Number_of_Shows',
                        action='store',
                        default=1,
                        widget='IntegerField',
                        help="Specify the number of shows happening at this venue")
    #Return parser function to main                    
    return parser.parse_args()