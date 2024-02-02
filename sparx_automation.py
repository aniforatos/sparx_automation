from win32com import client
import win32com.client
import pandas as pd
import xml.etree.ElementTree as ET
import logging
import webbrowser

class SparxAutomator:
    def __init__(self, file_path=None):
        self.log = logging.getLogger(self.__class__.__name__)

        if file_path is None:
            eaApp = win32com.client.Dispatch("EA.App")
            self.ea_repository = eaApp.Repository
        else:
            self.ea_repository = win32com.client.Dispatch("EA.Repository")
        
        self.log.info(f"Repository Connection Established")
        print(f"Repository Connection Established")
    
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

        # Extract column names from the first row
        column_names = [element.tag for element in root[0]]

        # Extract data from XML and create a list of dictionaries
        data = [{element.tag: element.text for element in row} for row in root]

        # Convert the list of dictionaries to a Pandas DataFrame
        df = pd.DataFrame(data, columns=column_names)
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

    def write_dataframe_series_to_html(self, series, file_name = None):
        

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

            # Write the HTML footer
            html_file.write('</body>\n</html>')
        
        # Open the file.
        webbrowser.open(file_name)