from win32com import client
import win32com.client
import pandas as pd
import xml.etree.ElementTree as ET
import logging
import webbrowser
import re

from jira import JIRA

class SparxAutomator:
    def __init__(self, file_path=None):
        self.log = logging.getLogger(self.__class__.__name__)

        if file_path is None:
            eaApp = win32com.client.Dispatch("EA.App")
            self.ea_repository = eaApp.Repository
        else:
            self.ea_repository = win32com.client.Dispatch("EA.Repository")
        
        # Load in the configuration file.
        self.config = ET.parse("./cfg/config.xml").getroot()
    
        self.log.info(f"Repository Connection Established")
        print(f"Repository Connection Established")
    

    def authenticate_jira(self):

        # Define login parameters.
        jira_config = self.config.find("jira")
        user_name = jira_config.find("email").text
        api_key = jira_config.find("api_key").text
        self.jira_url = 'https://resilienx.atlassian.net/'
        
        # Create JIRA object.
        self.jira = JIRA(self.jira_url, basic_auth=(user_name, api_key))
        
        # Check for successful authentication
        if self.jira.session is None:
            self.log.error("JIRA authentication failed.")
            print("Failed to authenticate with JIRA!")
            return False
        
        return True
        
    def add_comment_to_jira_story(self, story_id, html):
        self.jira.add_comment(story_id, html)

    def get_current_diagram_id(self):
        """Returns the ID of the Current Diagram.

        Returns:
            int: Diagram ID, could be None.
        """
        active_diagram = self.ea_repository.GetCurrentDiagram()

        # Log to inform current diagram.
        self.log.info(f"Current Diagram Name: {active_diagram.name}")
        self.log.info(f"Current Diagram ID: {active_diagram.DiagramID}")
        
        print(f"Current Diagram Name: {active_diagram.name}")
        is_correct = input("Is this the diagram you want?\nIf not, go and select the diagram in Sparx and return here.\nDiagram Correct (y/n) or Quit (q): ")
        if is_correct.lower() == "y":
            return active_diagram.DiagramID
        elif is_correct.lower() == "n":
            return self.get_current_diagram_id()
        else:
            return None
    
    def execute_sql_query(self, query):

        self.log.debug(f"Executing SQL Query: {query}")

        # Execute the SQL query
        result_set = self.ea_repository.SQLQuery(query)

        # Parse the XML data
        root = ET.fromstring(result_set)
        root = root.find("Dataset_0").find("Data")

        # Extract column names from the first row. With new EA, seems like the
        # SQL query names get capitalized, so I wll make every single column lowercase.
        column_names = [element.tag for element in root[0]]

        # Extract data from XML and create a list of dictionaries
        data = [{element.tag: element.text for element in row} for row in root]

        # Convert the list of dictionaries to a Pandas DataFrame
        df = pd.DataFrame(data, columns=column_names)

        # Convert all column names to lowercase
        df.columns = df.columns.str.lower()

        self.log.debug(f"Dataframe from SQL Query: {df}")

        return df

    def query_for_diagram_comments(self, diagram_id):
        q = f"""SELECT
                t_diagram.Name AS DiagramName,
                t_object.Name AS ElementName,
                t_object.Note AS Comment,
                t_object.Object_Type as Type
            FROM
                t_diagramobjects
            JOIN
                t_diagram ON t_diagram.Diagram_ID = t_diagramobjects.Diagram_ID
            JOIN
                t_object ON t_object.Object_ID = t_diagramobjects.Object_ID

            WHERE
                (t_object.Object_Type = 'Note' OR t_object.Stereotype = 'Note') AND
                t_diagram.Diagram_ID = {diagram_id};"""
                
        return self.execute_sql_query(q)

    def add_jira_comment(self, story_id, content):
        """Generates a comment in JIRA via the API.

        Args:
            story_id (string): Usually something along the lines of RCD-1
            content (string): The comment from the diagram highlighted.
        """

        # Add bullets instead of list tags
        content = content.replace('<li>', '* ')

        # Remove html tags.
        clean_text = re.sub('<.*?>', '', content)
        self.log.info(f"Adding comment: {clean_text}")

        # Add the comment.
        self.jira.add_comment(story_id, clean_text)

    def write_dataframe_series_to_html(self, series, story_id, file_name = None):
        

        # Replace 'output_file.txt' with your desired output file path
        if file_name is None:
            file_name = "output_file.html"

        self.log.info(f"Saving HTML to: {file_name}")

        # Extract the contents and send to list
        formatted_text_contents = series.tolist()
        
        # Write the contents to a rich-text HTML file
        with open(file_name, 'w', encoding='utf-8') as html_file:
            # Write the HTML header
            html_file.write('<html>\n<head></head>\n<body>\n')

            # Write each formatted text content            
            for content in formatted_text_contents:
                html_file.write(f'{content}<br><br>')
                try:
                    if story_id != "":
                        self.add_jira_comment(story_id, content)
                except Exception as e:
                    self.log.error(f"Error trying to add comment: {e}")
                    print("Comment add failure... See log.")
                
            # Write the HTML footer
            html_file.write('</body>\n</html>')
        
        if story_id == "":
            # Open the file.
            webbrowser.open(file_name)
        else:
            webbrowser.open(f"{self.jira_url}/browse/{story_id}")