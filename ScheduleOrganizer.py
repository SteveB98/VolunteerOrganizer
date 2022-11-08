from gooey import Gooey, GooeyParser
import os
import json

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
    parser.add_argument('Number of Volunteers',
                        action='store',
                        widget='IntegerField',
                        default=stored_args.get('Number_of_Volunteers'),
                        help="Specify the number of volunteers that will be present")
    #Venue Name
    parser.add_argument('Venue Name',
                        action='store',
                        default=stored_args.get('Venue_Name'),
                        help='Specify the name of the venue')
    #List of venue positions
    parser.add_argument('Volunteer/Supervisor Positions',
                        choices=['Beer Garden','Front of House','Green Team','Hospitality','Merchandise','Staging','Security'],
                        action='store',
                        default=stored_args.get('Volunteer/Supervisor Positons'),
                        widget='Listbox',
                        nargs="*",
                        metavar="Volunteer/Supervisor Positons",
                        help='Specify the work positons for this venue')
    #Type of supervisors at venue
    parser.add_argument('Volunteer/Supervisor Positions',
                        choices=['Beer Garden','Front of House','Green Team','Hospitality','Merchandise','Staging','Security'],
                        action='store',
                        default=stored_args.get('Volunteer/Supervisor Positons'),
                        widget='Listbox',
                        nargs="*",
                        metavar="Volunteer/Supervisor Positons",
                        help='Specify the work positons for this venue')
    parser.add_argument('Number of Shows',
                        action='store',
                        default=stored_args.get('Number_of_Shows'),
                        widget='IntegerField',
                        help="Specify the number of shows happening at this venue")
    args = parser.parse_args()

if __name__ == '__main__':
    conf = parse_args()
    print("Creating Venue Schedule Grid")
    #Pass all variables into excel spreadsheet generator program
    print("Done")