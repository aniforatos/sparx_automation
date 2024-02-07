# sparx_automation
A tool to be built on that will enable the automation of sparx.

This will create, and update comments on a specific JIRA story based on diagram comments as well.

## GUI
This tool now has a GUI associated with it, just run:
* `pip install -r requirements.txt`
*  `python ./main_ui.py`

Use the "Setup" tab to enter your JIRA login information.

This will give you the ability to set up your JIRA account and give you instant feedback as to what diagram you are focused on and the ability to changed your
issue ID on the fly.

## Comment Extraction
If you are looking to extract comments from a diagram in EA, there are a couple of things that need to happen first.
1. Enterprise Architect should be open (this will make life way easier)
2. Make sure that your diagram is in-view (i.e., The diagram is displaying in your EA Application)
    * The script will make you double check that the name of the diagram is the one you want.

## Running the Tool
simply run `python main.py -h` to understand the different options you have for automation.

