
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


def extract_all_child_requirements(sparx):
    logging.info("Extracting all child requirements from a user selected repository.")
    # Notify user to select the package they want
    input("Ensure you have the package selected in EA (Enter to Continure): ")

    # Get the currently selected package in the browser window.
    selected_package = sparx.ea_repository.GetTreeSelectedPackage()
    logging.info(f"The selected package is: {selected_package.Name}")
    # Capture a list of package objects
    package_list = sparx.get_child_packages(selected_package)

    # Query all requirements from the package list.
    df = sparx.query_requirements_from_package_list(sparx.package_list_to_ids(package_list))

    # Send it to excel.
    df.to_excel("child_requirements.xlsx")

def main():
    parser = argparse.ArgumentParser(description='Sparx Automation Tool', formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--model_path', type=str, help='Full path to the model file you want to open.\nOnly used if you want to open a new instance of Sparx (not recommended).')
    parser.add_argument('--action', required=True, type=str, choices=["c", "r", "rz", "cr"], help="Available options are: \n\tc: Extract comments from current diagram.\
                        \n\tr: Color requirements based on status\n\trv: Return requirement colors back to default\n\tcr: Extract all child requirements from a package.")
    
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
    elif args.action == "cr":
        extract_all_child_requirements(sparx)

main()