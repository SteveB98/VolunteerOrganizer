from gooey import Gooey, GooeyParser
import os
import mkl
import json
import openpyxl as pyxl

@Gooey(program_name="Venue Grid Generator")
def parse_args():
    """ Use GooeyParser to build up the arguments we will use in our script
    Save the arguments in a default json file so that we can retrieve them
    every time we run the script.
    """
    stored_args = {}
    # get the script name without the extension & use it to build up
    # the json filename
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    args_file = "{}-args.json".format(script_name)
    # Read in the prior arguments as a dictionary
    if os.path.isfile(args_file):
        with open(args_file) as data_file:
            stored_args = json.load(data_file)
    parser = GooeyParser(description='Produce a Venue Schedule')
    #Creating Dimensions for schedule grid
    #Number of volunteers at venue
    #Include functionality of selecting number of volunteer slots for each position
    #Include functionality to pick a selection of dates
    #Include functionality of specifying a shift time, for volunteers and supervisors
    parser.add_argument('--Number_of_Volunteers',
                        action='store',
                        widget='IntegerField',
                        default=1,
                        help="Specify the number of volunteers that will be present")
    #Venue Name
    parser.add_argument('Venue_Name',
                        action='store',
                        default=stored_args.get('Venue_Name'),
                        help='Specify the name of the venue')
    #List of venue positions
    parser.add_argument('Volunteer_Supervisor_Positions',
                        choices=['Beer Garden','Front of House','Green Team','Hospitality','Merchandise','Staging','Security'],
                        action='store',
                        default=stored_args.get('Volunteer/Supervisor Positons'),
                        widget='Listbox',
                        nargs="+",
                        metavar="Volunteer/Supervisor Positons",
                        help='Specify the work positons for this venue')
    #Type of supervisors at venue
    parser.add_argument('Supervisor_Positions',
                        choices=['Beer Garden','Front of House','Green Team','Hospitality','Merchandise','Staging','Security'],
                        action='store',
                        default=stored_args.get('Supervisor_Positons'),
                        widget='Listbox',
                        nargs="+",
                        metavar="Supervisor Positons",
                        help='Specify the work positons for this venue')
    parser.add_argument('--Number_of_Shows',
                        action='store',
                        default=1,
                        widget='IntegerField',
                        help="Specify the number of shows happening at this venue")
    #Return parser function to main                    
    return parser.parse_args()

if __name__ == '__main__':
    print("Creating Venue Schedule Grid")
    wb = pyxl.load_workbook("C:\\Users\\12508\\Documents\\ProgrammingStuff\\VolunteerOrganizer\\Test.xlsm", keep_vba=True)
    sheet = wb.active
    sheet.title = 'Hello World'
    wb.save("Test.xlsm")
    args = parse_args()
    print(args.Number_of_Volunteers)
    print(args.Venue_Name)
    print(args.Volunteer_Supervisor_Positions)
    print(args.Supervisor_Positions)
    print(args.Number_of_Shows)
    #Pass all variables into excel spreadsheet generator program

    #Make Venue Name Header

    #Generate columns based on number of performance days, length based on # volunteer positions and supervisor positions;
    #cross-examine Supervisor_Positions in cell exclusion case

    #Create cell formatting (size, colour, etc)

    #Insert volunteer positions into designated cell

    #Insert Dates into each cell


    sheet['A1'] = args.Number_of_Volunteers
    sheet['A2'] = args.Venue_Name
    sheet['A3'] = str(args.Volunteer_Supervisor_Positions)
    sheet['A4'] = str(args.Supervisor_Positions)
    sheet['A5'] = args.Number_of_Shows
    wb.save("Test.xlsm")
        
    print("Done")
