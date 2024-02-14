
import argparse
from sparx_automation import SparxAutomator
import logging
import logging.config

# Uncomment to disable logging outside of this application.
logging.config.dictConfig({
    'version': 1,
    'disable_existing_loggers': True,
})

logging.basicConfig(filename="sparx_automation.log", filemode="w", format='%(asctime)s [%(name)s.%(funcName)s] - %(levelname)s - %(message)s',
                    level=logging.DEBUG)

def color_requirements_by_status(sparx, revert=False):
    """This kicks off the automation process that is used to color requirements
    on a given diagram by its Status value.

    Args:
        sparx (EAObject): This is the COM object that represents the enterprise architect instance.
        revert (bool, optional): This allows the user to put the colors back to their default state. Defaults to False.
    """

    # Get the current diagram object.
    d_obj = sparx.get_current_diagram_name()

    # Make sure there is an active diagram object.
    if d_obj.DiagramID is None:
        logging.error("User aborted the diagram ID selection. Diagram ID is None")
        return

    # Query for the diagram object requirements present.
    df = sparx.query_for_diagram_requirements(d_obj.DiagramID)
    
    # Loop through the diagram objects and color them appropriately
    sparx.loop_diagram_objects(d_obj, df, revert)

def extract_comments_from_diagram(sparx):
    jira_suc = sparx.authenticate_jira()

    d_obj = sparx.get_current_diagram_name()
    d_id = d_obj.DiagramID

    if d_id is None:
        logging.error("User aborted the diagram ID selection. Diagram ID is None")
        return
        
    if jira_suc:
        df = sparx.query_for_diagram_comments(d_id)
        story_id = input("Input JIRA Issue ID (e.g., RCD-1. Enter to skip): ")
    
    sparx.write_dataframe_to_html_and_jira(df, story_id)   

def main():
    parser = argparse.ArgumentParser(description='Sparx Automation Tool', formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--model_path', type=str, help='Only used if you want to open a new instance of Sparx (not recommended). \
                         Full path to the model file you want to open.')
    parser.add_argument('--action', required=True, type=str, choices=["c", "r", "rz"], help="Available options are: \n\tc: Extract comments from current diagram.\
                        \n\tr: Color requirements based on status\n\trv: Return requirement colors back to default")
    
    # Parse the command-line arguments
    args = parser.parse_args()

    # Define arguments
    model_path = args.model_path

    sparx = SparxAutomator(file_path=model_path)
    
    if args.action == "c":
        logging.info("Extracting comments from the current diagram...")
        extract_comments_from_diagram(sparx)
    elif args.action == "r":
        logging.info("Setting Element Colors on Requirements diagram based on status.")
        color_requirements_by_status(sparx)
    elif args.action == "rz":
        color_requirements_by_status(sparx, revert=True)

main()