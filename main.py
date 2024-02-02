
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

def extract_comments_from_diagram(sparx):
    d_id = sparx.get_current_diagram_id()
    
    if d_id is None:
        logging.error("User aborted the diagram ID selection. Diagram ID is None")
        return
    
    df = sparx.query_for_diagram_comments(d_id)
    sparx.write_dataframe_series_to_html(df["Comment"])   

def main():
    parser = argparse.ArgumentParser(description='Sparx Automation Tool', formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--model_path', type=str, help='Only used if you want to open a new instance of Sparx (not recommended). \
                         Full path to the model file you want to open.')
    parser.add_argument('--action', required=True, type=str, choices=["c"], help="Available options are: \n\tc: Extract Comments from Current Diagram.")
    
    # Parse the command-line arguments
    args = parser.parse_args()

    # Define arguments
    model_path = args.model_path

    sparx = SparxAutomator(file_path=model_path)

    if args.action == "c":
        logging.info("Extracting comments from the current diagram...")
        extract_comments_from_diagram(sparx)

main()