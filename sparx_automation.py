from win32com import client
import win32com.client

import pandas as pd
# disable chained assignments
pd.options.mode.chained_assignment = None 

import xml.etree.ElementTree as ET
import logging
import webbrowser
import re
import time
from tqdm import tqdm

from jira import JIRA

class SparxAutomator:
    def __init__(self, file_path=None):

        # Init logger.
        self.log = logging.getLogger(self.__class__.__name__)

        # Try to get current EA session or open a new one.
        if file_path is None:
            eaApp = win32com.client.Dispatch("EA.App")
            self.ea_repository = eaApp.Repository
        else:
            self.ea_repository = win32com.client.Dispatch("EA.Repository") 

        self.load_config()
        self.define_color_mapping()
        self.log.info(f"Repository Connection Established")
        print(f"Repository Connection Established")
    
    def load_config(self):
        # Load in the configuration file.
        self.config = ET.parse("./cfg/config.xml").getroot()

    def define_color_mapping(self):
        # Load the color mapping for requirements from config.
        req_color_config = self.config.find("requirement_colors")
        self.status_color_map = {
            "Proposed": eval(req_color_config.find("Proposed").text)[::-1],
            "Implemented": eval(req_color_config.find("Implemented").text)[::-1],
            "Approved": eval(req_color_config.find("Implemented").text)[::-1],
            "Default": (0, 0, 0)
        } 

    def check_jira_authentication(self):
        # Check for successful authentication
        try:
            curr_user = self.jira.current_user()
            self.log.debug(f"Current user information: {curr_user}")
        except:
            self.log.error("JIRA authentication failed.")
            print("Failed to authenticate with JIRA!")
            return False
        
        return True

    def authenticate_jira(self):
        """Authenticates via user email and an API key they generate for their email.

        Returns:
            boolean: Indication of whether or not the user is authenticated.
        """
        # Define login parameters.
        jira_config = self.config.find("jira")
        user_name = jira_config.find("acctEmail").text
        api_key = jira_config.find("jiraApiKey").text
        self.jira_url = 'https://resilienx.atlassian.net/'
        
        # Create JIRA object.
        self.jira = JIRA(self.jira_url, basic_auth=(user_name, api_key))
        
        # Check for auth success.
        return self.check_jira_authentication()

    def user_accept_current_diagram(self, active_diagram):
        # Prompt the user to check if the name is correct and what they want.
        is_correct = input("Is this the diagram you want?\nIf not, go and select the diagram in Sparx and return here.\nDiagram Correct (y/n) or Quit (q): ")
        if is_correct.lower() == "y":
            return active_diagram
        
        # If not, then re-run the function.
        elif is_correct.lower() == "n":
            return self.get_current_diagram_name()
        
        # Otherwise, quit the script.
        else:
            return None
    
    def loop_diagram_objects(self, d_obj, df, revert):
        for obj in tqdm(d_obj.DiagramObjects):
            if obj.ElementID in df["objectid"].values:                
                row = df.loc[df["objectid"] == obj.ElementID, :].iloc[0]               
                color = self.status_color_map[row["status"]]
                if revert:
                    color = self.status_color_map["Default"]
                
                # Log the color settings.
                self.log.debug(f"Setting ({row['elementname']}) to RGB color: {color[::-1]}")                

                rgb_integer = (color[0] << 16) + (color[1] << 8) + color[2]
                obj.SetStyleEx("BCol",str(rgb_integer))
                obj.Update()
                time.sleep(0.5)

    def get_current_diagram_name(self, input=True):
        """Returns the ID of the Current Diagram.

        Returns:
            int: Diagram ID, could be None.
        """
        try:

            # Get the active diagram.
            active_diagram = self.ea_repository.GetCurrentDiagram()

            # Log to inform current diagram.
            self.log.info(f"Current Diagram Name: {active_diagram.name}")
            self.log.info(f"Current Diagram ID: {active_diagram.DiagramID}")
            
            print(f"Current Diagram Name: {active_diagram.name}")
            if input:
                return self.user_accept_current_diagram(active_diagram)
            
            return active_diagram
            
        except AttributeError:
            self.log.error("No diagram active!")
            print("No diagram selected! Go select in Enterprise Architect")
            return None
        
    def execute_sql_query(self, query):
        """Execute a SQL query in Enterprise Architect.

        Args:
            query (str): The SQL query to capture data from EA.

        Returns:
            pandas.core.DataFrame: The dataframe that represents the results of the query.
        """
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
        df[["JiraCommentID", "JIRA_Task"]] = None
        df["objectid"] = df["objectid"].astype(int)

        self.log.debug(f"Dataframe from SQL Query: {df}")

        return df

    def query_for_diagram_requirements(self, diagram_id):
        sql_query = f"""SELECT
                t_diagram.Name AS DiagramName,
                t_object.Name AS ElementName,
                 t_object.Status as Status,
                t_object.Object_Type as Type,
                t_object.Object_ID as ObjectID               
            FROM
                t_diagramobjects
            JOIN
                t_diagram ON t_diagram.Diagram_ID = t_diagramobjects.Diagram_ID
            JOIN
                t_object ON t_object.Object_ID = t_diagramobjects.Object_ID 
        WHERE
            /* Specify the Diagram GUI ID */
            t_diagram.Diagram_ID = {diagram_id}
            AND t_object.Object_Type = 'Requirement';"""
        
        return self.execute_sql_query(sql_query)

    def query_for_diagram_comments(self, diagram_id):
        q = f"""SELECT
                t_diagram.Name AS DiagramName,
                t_object.Name AS ElementName,
                t_object.Note AS Comment,
                t_object.Object_Type as Type,
                t_object.Object_ID as ObjectID,
                t2.Name as Target_Name,
	            t_connector.End_Object_ID as TargetObject

            FROM
                t_diagramobjects
            JOIN
                t_diagram ON t_diagram.Diagram_ID = t_diagramobjects.Diagram_ID
            JOIN
                t_object ON t_object.Object_ID = t_diagramobjects.Object_ID            
            LEFT JOIN 
	            t_connector ON t_object.object_id = t_connector.Start_Object_ID
            LEFT JOIN
	            t_object t2 ON t_connector.End_Object_ID = t2.Object_ID

            WHERE
                (t_object.Object_Type = 'Note' OR t_object.Stereotype = 'Note') AND
                t_diagram.Diagram_ID = {diagram_id};"""
        try:
            df = self.execute_sql_query(q)

            # House cleaning to collapse dataframe duplicated OBJECTIDs into single row.
            df = df.groupby("objectid").agg(lambda x: x.tolist() if x.nunique() > 1 else x.iloc[0]).reset_index()
            
            return df
        except Exception as e:

            self.log.error("Failed to execute SQL Query, returning empty dataframe.")
            self.log.error(e)
            return pd.DataFrame(data={"Empty": [None]})

    def create_content(self, row):
        """Takes in the row to be added as a comment to JIRA and creates
        a comment body from multiple columns.

        Args:
            row (pandas.core.Series): Row containing information to populate the comment.

        Returns:
            str: Comment body string to be added to JIRA.
        """
        # Add the targets to the beginning of the comment.
        if type(row["target_name"]) == list:
            content = ', '.join(row["target_name"]) + "\n"
        elif type(row["target_name"]) == str:
            content = row["target_name"] + "\n"
        else:
            content = ""

        # Append the comment text to the target name header that was added.
        content += row["comment"]

        return content

    def add_jira_comment(self, story_id, row):
        """Generates a comment in JIRA via the API.

        Args:
            story_id (string): Usually something along the lines of RCD-1
            content (string): The comment from the diagram highlighted.
        """

        # Create content for JIRA comment
        content = self.create_content(row)

        # Add bullets instead of list tags
        content = content.replace('<li>', '* ')

        # Remove html tags.
        clean_text = re.sub('<.*?>', '', content)
        self.log.info(f"Adding comment: {clean_text}")
        
        # Add the comment.
        comment = self.jira.add_comment(story_id, clean_text)

        # Return the comment to be added to the dataframe.
        return comment.id
    
    def update_jira_comment(self, data_df_row, story_id, row):
                
        # Create content for JIRA comment
        comment_txt = self.create_content(row)

        # Get the comment
        comment = self.jira.comment(story_id, str(data_df_row.loc[0, "JiraCommentID"]))
        
        # Update the comment
        self.log.info(f"Update comment ID: {comment.id}")
        comment.update(body=comment_txt)

    def extract_comments_from_diagram(self, criteria_dict, html=False):
        d_id = str(criteria_dict["tcDiagramId"])

        if d_id is None:
            logging.error("User aborted the diagram ID selection. Diagram ID is None")
            return

        df = self.query_for_diagram_comments(d_id)
        if html:
            self.write_dataframe_to_html_and_jira(df, "")
        return df

    def write_dataframe_to_html_and_jira(self, df, story_id, file_name = None):
        """Writes a dataframe to HTML but also adds JIRA comments if the user specified a story for it.

        Args:
            df (pandas.core.DataFrame): Dataframe containing the notes pulled from the diagram.
            story_id (str): The JIRA story ID (e.g., RCD-1)
            file_name (str, optional): The name of the file to save the html to. Defaults to None.
        """
        # Replace 'output_file.txt' with your desired output file path
        if file_name is None:
            file_name = "output_file.html"

        # Load the comment database file.
        data_df = pd.read_csv("./data/comment_dataframe.csv", index_col=0)

        # Write the contents to a rich-text HTML file
        self.log.info(f"Saving HTML to: {file_name}")
        with open(file_name, 'w', encoding='utf-8') as html_file:
            # Write the HTML header
            html_file.write('<html>\n<head></head>\n<body>\n')

            # Write each formatted text content            
            for i in range(df.shape[0]):
                row = df.iloc[i]
                html_file.write(f'{row["comment"]}<br><br>')
                try:
                    if story_id != "":
                        
                        # Check if the object ID exists in the database.
                        if row["objectid"] in data_df["objectid"].values and story_id in data_df["JIRA_Task"].values:

                            # Define the sub-dataframe filtered on object ID.
                            data_df_row = data_df.loc[data_df['objectid'] == row['objectid'], :].reset_index()

                            # Check if the comment has changed.
                            if row["comment"] == data_df_row.loc[0, "comment"]:

                                # For a comment that is unchanged, do nothing.
                                self.log.info(f"JIRA Comment {row['JiraCommentID']} already exists.")
                            
                            # Otherwise, Update the JIRA comment in JIRA and in the dataframe.
                            else:

                                # Update in JIRA.
                                self.log.info(f"Updating JIRA Comment: {row['JiraCommentID']}")
                                self.update_jira_comment(data_df_row, story_id, row)

                                # Update dataframe.
                                data_df.loc[data_df['objectid'] == row['objectid'], "comment"] = row["comment"]
                        
                        # If the object ID doesnt exist, its a new diagram comment.
                        else:

                            # Add the JIRA comment to JIRA
                            self.log.info("Adding new JIRA comment.")
                            c_id = self.add_jira_comment(story_id, row)   

                            # Add the new comment as a row in the database.
                            row.loc["JiraCommentID"] = int(c_id); row.loc["JIRA_Task"] = story_id
                            data_df.loc[len(data_df)] = row

                except Exception as e:
                    self.log.error(f"Error trying to add comment: {e}")
                    print("Comment add failure... See log.")
                
            # Write the HTML footer
            html_file.write('</body>\n</html>')
        
        # Re-write the database file.
        data_df.to_csv("./data/comment_dataframe.csv")
        
        # Open the comments in JIRA or the saved HTML file.
        if story_id == "":
            # Open the file.
            webbrowser.open(file_name)
        else:
            webbrowser.open(f"{self.jira_url}/browse/{story_id}")
        
        # Return the comments that are currently stored in JIRA
        return data_df[(data_df["objectid"].isin(df["objectid"])) & (data_df["JIRA_Task"] == story_id)]

    def manage_status(self, resource):
        try:
            if resource == "progress":
                return ""
        except:
            self.log.exception("message")

class ConfigMgr:
    def __init__(self) -> None:
        self.tree = ET.parse("./cfg/config.xml")
        self.root = self.tree.getroot()
        self.comment_extraction = self.root.find("comment_extraction")
        self.jira_cfg = self.root.find("jira")
    
    def pull_criteria(self, field, type):        
        return self.root.find(type).find(field).text
    
    def store_criteria(self, criteria_dict, type=None):
        try:
            cfg = self.root.find(type)
        except Exception as e:
            print(e)
            return False
        
        for key in list(criteria_dict.keys()):
            cfg.find(key).text = str(criteria_dict[key])
        
        self.tree.write("./cfg/config.xml")
    
    def assign_type(self, criteria_dict):
        for key in list(criteria_dict.keys()):
            if key == "tcIssueId":
                criteria_dict[key] = str(criteria_dict[key])
        
        return criteria_dict