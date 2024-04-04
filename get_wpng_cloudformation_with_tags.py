import subprocess
import json
from openpyxl import Workbook

# AWS CLI command to describe CloudFormation stacks
aws_command = "aws cloudformation describe-stacks --region us-east-1 --query 'Stacks[?Tags[?Key==`AppID` && Value==`APP-2313`]]'"

# Selected tag keys
selected_tag_keys = ["appCode", "envName"]  # Add or remove tag keys as needed

# Selected parameters
selected_parameter_keys = ["EnvName", "ServiceName", "ComponentName", "MaxInstances", "MinInstances", "InstanceType", "AppVersion", "AMI", "InstanceName"]  # add or remove parameters as needed

try:

    # Execute AWS CLI command and capture output
    output = subprocess.check_output(aws_command, shell=True)
    stacks = json.loads(output)

    # Create a new Excel workbook
    wb = Workbook()
    ws = wb.active

    # Write header row
    ws.append(["Stack Name", "Stack Status", "Creation Time", "appCode-TAG", "envName-TAG", "EnvName", "ServiceName", "ComponentName", "MaxInstances", "MinInstances", "InstanceType", "AppVersion", "AMI", "InstanceName"])
    print("header written to cloudformation_stacks.xlsx successfully.")

    # Iterate over stacks and write data to Excel
    for stack in stacks:
        stack_name = stack["StackName"]
        stack_status = stack["StackStatus"]
        creation_time = stack["CreationTime"]
    
        tags = {tag["Key"]: tag["Value"] for tag in stack.get("Tags", [])}  # Extract tags
        Parameter = {Parameter["ParameterKey"]: Parameter["ParameterValue"] for Parameter in stack.get("Parameters", [])}  # Extract parameters

        # Extract values for selected tag keys
        selected_tag_values = [tags.get(key, "") for key in selected_tag_keys]

        # Extract values for selected tag keys
        selected_parameter_values = [Parameter.get(key, "") for key in selected_parameter_keys]


        # ws.append([stack_name, stack_status, creation_time])
       # ws.append([stack_name, stack_status, creation_time, tags])
        ws.append([stack_name, stack_status, creation_time] + selected_tag_values + selected_parameter_values)

    # Save the Excel file
    wb.save("cloudformation_stacks.xlsx")
    print("Data written to cloudformation_stacks.xlsx successfully.")
    

except subprocess.CalledProcessError as e:
    print("Error executing AWS CLI command:", e)
except json.JSONDecodeError as e:
    print("Error decoding JSON:", e)
except Exception as e:
    print("An unexpected error occurred:", e)
