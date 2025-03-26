import argparse
import json
import os
import re
import tempfile
import time as system_time
import urllib.parse
import xml.etree.ElementTree as ET
from connection import login_to_tableau
import openpyxl
import pandas as pd
import requests
import tableauserverclient as TSC
# import openpyxl
from openpyxl.styles import Font
from tableaudocumentapi import Workbook
# from msystem import sys_variables
from tableauserverclient import *

from config_reader.reader_factory import ReaderFactory
from logger.logger import *

base_url = ""
cloud_url = ""

workbooksname = []
log = []
migrate = []
description_status = []
customname = []
custom_log = []
custom_migrate = []


def get_project_id(projects, project_path, server):
    # Base case: If the project path is empty, return None
    if not project_path:
        return None

    current_project = None
    for project in projects:
        # Case-insensitive comparison
        if project.name.lower() == project_path[1].lower():
            current_project = project
            break

    if current_project is None:
        return None

    for part in project_path[2:]:
        child_projects = [p for p in TSC.Pager(
            server.projects) if p.parent_id == current_project.id]
        current_project = None
        for child in child_projects:
            if child.name.lower() == part.lower():  # Case-insensitive comparison
                current_project = child
                break

        if current_project is None:
            return None

    return current_project.id


def get_project(cloud_server, project_name):
    """
    Get a top-level project by its name, ignoring case.
    :param cloud_server: Tableau Server or Tableau Online instance
    :param project_name: Name of the top-level project
    :return: Project object if found, else None
    """
    all_projects, _ = cloud_server.projects.get()
    for project in all_projects:
        if project.name.lower() == project_name.lower():
            return project
    return None


def get_subproject(cloud_server, parent_project_id, subproject_name):
    """
    Get a subproject by its name and parent project ID, ignoring case.
    :param cloud_server: Tableau Server or Tableau Online instance
    :param parent_project_id: ID of the parent project
    :param subproject_name: Name of the subproject
    :return: Subproject object if found, else None
    """
    all_projects, _ = cloud_server.projects.get()
    for project in all_projects:
        if project.parent_id == parent_project_id and project.name.lower() == subproject_name.lower():
            return project
    return None


def create_temp_dir(user_path=None):
    """
    Create a temporary directory in the specified path or the current directory.

    Args:
    user_path (str): The directory path where the temporary directory should be created.
                     If not specified, the current directory is used.

    Returns:
    str: The path to the created temporary directory.
    """
    # If user_path is provided, use it; otherwise, use the current working directory
    if user_path:
        # Ensure the path is absolute
        base_path = os.path.abspath(user_path)
    else:
        base_path = os.getcwd()

    # Ensure the base path exists
    if not os.path.exists(base_path):
        logger_error(f"The specified path does not exist: {base_path}")
        exit(1)

    # Check if we have permission to write in the directory
    if not os.access(base_path, os.W_OK):
        logger_error(f"Write permission denied for the path: {base_path}")
        exit(1)

    # Create a temporary directory within the specified path
    try:
        temp_dir = tempfile.mkdtemp(dir=base_path)
    except PermissionError as e:
        logger_error(f"Failed to create a temporary directory: {e}")
        exit(1)

    return temp_dir


def get_full_project_path(all_projects, project_id):
    # Create a dictionary for quick lookup of projects by their ID
    project_lookup = {project.id: project for project in all_projects}
    path = []
    current_id = project_id

    while current_id:
        project = project_lookup.get(current_id)
        if not project:
            break
        project_name = project.name
        project_name = project_name.replace('/', '_fwd_SLASH_')
        path.append(project_name)
        current_id = project.parent_id

    # Reverse the path to get the full path from root to target project
    full_path = '/'.join(reversed(path))

    # Ensure the path starts with a '/'
    if not full_path.startswith('/'):
        full_path = '/' + full_path

    return full_path if full_path else 'Top-level-project'


def get_child_project_id(projects, project_path, server):
    # Base case: If the project path is empty, return None
    if not project_path:
        return None
    for project in projects:
        project_path_ = project_path[0].replace('_fwd_SLASH_', '/')
        if project.name.lower() == project_path_.lower():
            if len(project_path) == 1:
                return project.id
            else:
                child_projects = [p for p in TSC.Pager(
                    server.projects) if p.parent_id == project.id]
                child_project_id = get_child_project_id(
                    child_projects, project_path[1:], server)
                if child_project_id:
                    return child_project_id
    return None


def get_project_ids(path, server):
    all_projects = list(TSC.Pager(server.projects))
    main_path = path[1:]
    if '/' in main_path:
        path_parts = main_path.split('/')
        return get_child_project_id(all_projects, path_parts, server)
    else:
        for pro in all_projects:
            if pro.name.lower() == main_path.lower() and pro.parent_id is None:
                return pro.id
    return None


def download(server, project_paths):
    try:
        # Create a temporary directory for downloads
        temp_dir = create_temp_dir()
        # create temp directory to download custom view
        custom_temp_dir = create_temp_dir()
        tableau_server, credentials = login_to_tableau(
            server["credentials"]["server"])

        # Authenticate the server
        with tableau_server.auth.sign_in(credentials):
            server_project = list(TSC.Pager(tableau_server.projects))
            server_custom = list(TSC.Pager(tableau_server.custom_views))
            workbook_details = []
            description = []
            for project_path in project_paths:
                # Split the project path into components
                project_path_components = project_path.strip('/').split('/')

                # Initialize variables to track the current project hierarchy
                current_parent_id = None
                current_project = None

                current_project = get_project_ids(project_path, tableau_server)

                project_workbook_items = [
                    wb for wb in TSC.Pager(tableau_server.workbooks) if wb.project_id == current_project
                ]
                names = [workbook.name for workbook in project_workbook_items]
                logger_info(f"Downloaded workbooks: {names}")
                logger_info(
                    f"Total number of workbooks in project '{project_path}': {len(names)}")

                # Download the workbooks in the temp folder
                for workbook in project_workbook_items:
                    # Construct the folder path for the project hierarchy
                    encoded_path_components = [urllib.parse.quote(
                        component) for component in project_path_components]
                    sanitized_path_components = [sanitize_filename(
                        component) for component in encoded_path_components]
                    project_folder = os.path.join(
                        temp_dir, *sanitized_path_components)

                    # Create folders for the project hierarchy if they don't exist
                    if not os.path.exists(project_folder):
                        os.makedirs(project_folder)

                    workbook.name = workbook.name.replace('/', '_fwd_SLASH_')
                    # Define the file path for the workbook
                    encoded_workbook_name = urllib.parse.quote(workbook.name)
                    sanitized_datasource_name = sanitize_filename(
                        encoded_workbook_name)
                    file_path = os.path.join(
                        project_folder, sanitized_datasource_name)
                    try:
                        workbook = tableau_server.workbooks.get_by_id(
                            workbook.id)

                        tableau_server.workbooks.populate_connections(workbook)
                        if workbook.id != '':
                            for connection_w in workbook.connections:

                                datasource = tableau_server.datasources.get_by_id(
                                    connection_w.datasource_id)
                                datasource_projectpath = get_full_project_path(
                                    TSC.Pager(tableau_server.projects), datasource.project_id)
                                workbook_details.append({
                                    "datasource_id": datasource.id,
                                    "datasource_name": datasource.name,
                                    "datasource_projectpath": datasource_projectpath,
                                    "filepath": file_path,
                                    "workbook_name": workbook.name,
                                    "server_address": connection_w.server_address,
                                    "server_port": connection_w.server_port,
                                    "connection_type": connection_w.connection_type,
                                    "connection_username": connection_w.username,
                                    "connection_password": connection_w.password,
                                    "connection_embed_password": connection_w.embed_password,
                                })
                            description.append(
                                {"filepath": file_path, "description": workbook.description, "workbook_id": workbook.id})
                        download_customview(
                            tableau_server, server_custom, workbook.id, custom_temp_dir)
                        tableau_server.workbooks.download(
                            workbook.id, filepath=file_path)
                        logger_info(
                            f"Downloaded Workbook '{workbook.name}' to '{file_path}'")
                    except Exception as e:
                        logger_error(
                            f"Failed to download workbook '{workbook.name}': {e}")

        # exit()
        # Return the path to the downloaded workbooks
        cloud_workbooks = temp_dir
        return cloud_workbooks, workbook_details, description, custom_temp_dir
    except Exception as download_error:
        logger_error(
            f"An error occurred during workbook download: {download_error}")
        return None  # Indicate failure by returning None


def sanitize_filename(filename):
    """     Remove or replace characters that are invalid for filenames on Windows.     """
    # Replace invalid characters with an underscore
    return re.sub(r'[<>:"/\\|?*]', '_', filename)


def convert_path_format(mixed_path):
    # Convert the mixed path to a consistent format
    normalized_path = os.path.normpath(mixed_path)

    # Replace any forward slashes with backslashes
    normalized_path = normalized_path.replace('/', '\\')

    # Escape the backslashes
    escaped_path = normalized_path.replace('\\', '\\\\')

    # Decode URL-encoded parts of the path
    decoded_path = urllib.parse.unquote(escaped_path)

    # Remove the .twbx extension
    if decoded_path.endswith('.twbx'):
        decoded_path = decoded_path[:-5]
    elif decoded_path.endswith('.twb'):
        decoded_path = decoded_path[:-4]

    return decoded_path


def download_customview(source_server, all_custom_view, workbook_id, custom_temp_dir):
    # Setting SSL value as it is set in config file
    verify_ssl = source_server._http_options.get('verify', True)
    for custom_view in all_custom_view:
        if custom_view.workbook.id == workbook_id:
            workbook_dir = os.path.join(custom_temp_dir, str(workbook_id))
            os.makedirs(workbook_dir, exist_ok=True)
            headers = {
                # Assuming 'source_server' has 'auth_token' attribute
                'X-Tableau-Auth': source_server.auth_token
            }

            # # Dynamically setting the API version from source_server
            # source_api_version = source_server.version
            # url = f"{source_server.baseurl.replace('/api/' + source_api_version, '/api/exp')}/sites/{source_server.site_id}/customviews/{custom_view.id}/content"
            # logger_info(f"Source url: {url}")

            url = f"{source_server.baseurl.replace('/api/3.21', '/api/exp')}/sites/{source_server.site_id}/customviews/{custom_view.id}/content"

            response = requests.get(url, headers=headers, verify=verify_ssl)

            get_custom_view = source_server.custom_views.get_by_id(
                custom_view.id)
            # Check if the request was successful
            if response.status_code == 200:
                if get_custom_view.shared == "True":
                    shared = 'T'
                else:
                    shared = 'F'
                custom_view.name = custom_view.name.replace('/', '_fwd_SLASH_')
                encoded_workbook_name = urllib.parse.quote(custom_view.name)
                sanitized_customview_name = sanitize_filename(
                    encoded_workbook_name)
                file_path = os.path.join(
                    workbook_dir, f'{sanitized_customview_name}_&_{custom_view.owner.name}_&_{shared}')
                # Save the response content to a file
                with open(file_path, 'wb') as file:  # Using custom_view.name
                    file.write(response.content)
                logger_info("Custom view downloaded successfully.")
            else:
                logger_error(
                    f"Failed to download custom view. Status code: {response.status_code}")
                logger_error(f"Error message: {response.text}")


def publish_custom_view(cloud_server, descriptions, user_list, custom_temp_dir, published_workbook_id, cloud_allusers):
    try:
        # Extract workbook IDs and corresponding published workbook IDs from descriptions
        workbook_id_map = {desc['workbook_id']: desc.get(
            'published_workbook_id') for desc in descriptions}

        # Get SSL verification setting from server config
        verify_ssl = cloud_server._http_options.get('verify', True)
        # Process directories that match workbook IDs
        for workbook_id, pub_workbook_id in workbook_id_map.items():
            if pub_workbook_id == published_workbook_id:
                dir_path = os.path.join(custom_temp_dir, workbook_id)
                if os.path.exists(dir_path) and os.path.isdir(dir_path):
                    # Process each file in the matched directory
                    for file in os.listdir(dir_path):
                        file_path = os.path.join(dir_path, file)
                        # Extract custom_view name and owner name from file name
                        if '_&_' in file:
                            custom_view_name, owner_name, shared = file.split(
                                '_&_')
                        else:
                            custom_view_name = file
                            owner_name = "Unknown"
                        user_id = ''
                        user_name = ''
                        for filter_user in user_list:
                            if owner_name == filter_user[1]:
                                user_id = next(
                                    (user.id for user in cloud_allusers if user.email == filter_user[2]), None)
                                user_name = next(
                                    (user.name for user in cloud_allusers if user.email == filter_user[2]), None)
                                break
                        decoded_file_name = urllib.parse.unquote(
                            custom_view_name)
                        sanitized_file_name = sanitize_filename(
                            decoded_file_name)
                        decoded_file_name = decoded_file_name.replace(
                            '_fwd_SLASH_', '/')
                        decoded_file_name = decoded_file_name.replace(
                            '&', '&amp;')

                        # Dynamically get the API version from cloud_server
                        # cloud_api_version = cloud_server.version
                        # url = f"{cloud_server.baseurl.replace('/api/' + cloud_api_version, '/api/exp')}/sites/{cloud_server.site_id}/customviews?overwrite=True"
                        # logger_info(f"Cloud url: {url}")

                        url = f"{cloud_server.baseurl.replace('/api/3.21', '/api/exp')}/sites/{cloud_server.site_id}/customviews?overwrite=True"

                        boundary_string = "my_boundary_string"  # Unique boundary string
                        # Create the multipart/mixed body content
                        body = (
                            f"--{boundary_string}\r\n"
                            f"Content-Disposition: name=\"request_payload\"\r\n"
                            f"Content-Type: text/xml\r\n\r\n"
                            f"<tsRequest>\r\n"
                            f"    <customView name=\"{decoded_file_name}\" shared=\"{shared}\">\r\n"
                            f"        <workbook id=\"{published_workbook_id}\"/>\r\n"
                            f"        <owner id=\"{user_id}\" />\r\n"
                            f"    </customView>\r\n"
                            f"</tsRequest>\r\n"
                            f"--{boundary_string}\r\n"
                            f"Content-Disposition: name=\"tableau_customview\"; filename=\"{file_path.split('/')[-1]}\"\r\n"
                            f"Content-Type: image/png\r\n\r\n"
                        )
                        # Read the content of the custom view file
                        with open(file_path, 'rb') as f:
                            file_content = f.read()
                        # Append file content and the final boundary string
                        body += file_content.decode('latin-1') + \
                            f"\r\n--{boundary_string}--\r\n"

                        headers = {
                            "X-Tableau-Auth": cloud_server.auth_token,
                            "Content-Type": f"multipart/mixed; boundary={boundary_string}"
                        }

                        response = requests.post(
                            url, headers=headers, data=body.encode('latin-1'), verify=verify_ssl)
                        decoded_file_name = decoded_file_name.replace(
                            '&amp;', '&')
                        customname.append(decoded_file_name)
                        # print("response.status_code",response.status_code)
                        # print("response.text",response.text)
                        if response.status_code == 200:
                            logger_info(
                                f"\t\tCustom View {decoded_file_name} published successfully")
                            # Delete the file after successful publishing
                            os.remove(file_path)
                            custom_migrate.append('Migrated Custom View')
                            custom_log.append(
                                f"Custom View {decoded_file_name} published successfully")
                        else:
                            logger_error(
                                f"Failed to publish custom view. Status code: {response.status_code}")
                            logger_error(f"Error message: {response.text}")
                            custom_migrate.append(
                                'Failed to Migrate Custom View')
                            custom_log.append(
                                f"Failed to Migrate Custom View: {response.text}")
        log_data = pd.DataFrame({
            'Custom View Name': customname,
            'Log': custom_log,
            'Status': custom_migrate,
        })
        file_name = 'migration_logs.xlsx'
        sheet_name = 'custom view'
        current_directory = os.getcwd()
        file_path = os.path.join(current_directory, file_name)
        try:
            workbook = openpyxl.load_workbook(file_path)
        except FileNotFoundError:
            workbook = openpyxl.Workbook()
            default_sheet = workbook["Sheet"]
            workbook.remove(default_sheet)
        if sheet_name not in workbook.sheetnames:
            workbook.create_sheet(sheet_name)
        sheet = workbook[sheet_name]
        sheet.delete_rows(1, sheet.max_row)
        sheet.append(log_data.columns.tolist())
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        for row in log_data.itertuples(index=False, name=None):
            sheet.append(row)
        workbook.save(file_path)
    except Exception as e:
        logger_error(f"An unexpected error occurred: {e}")


def migrate_workbooks(cloud_server, workbooks_dir, paths, workbook_details, descriptions, user_list, custom_temp_dir):
    try:
        logger_info("------Starting publishing specific workbooks------")
        # Get SSL verification setting from server config
        verify_ssl = cloud_server._http_options.get('verify', True)

        total_specific_workbooks = 0  # Initialize counter for specific workbooks processed
        # Initialize counter for specific workbooks successfully published
        total_specific_published = 0
        total_specific_failed = 0  # Initialize counter for specific workbooks failed to publish

        failed_workbooks = []  # List to store names of failed workbooks

        # Ensure paths is a list of lists
        if not all(isinstance(entry, list) and len(entry) == 2 for entry in paths):
            logger_error(
                "Invalid format for paths. Expected a list of lists with two elements each.")
            return

        mapping_file = "datasource_path.xlsx"
        df = pd.read_excel(mapping_file)
        cloud_users = list(TSC.Pager(cloud_server.users))
        # Iterate through the list and publish each workbook
        for path_entry in paths:
            source_path, destination_path = path_entry

            # Combine the base directory with the relative source path
            encoded_source_path = urllib.parse.quote(
                source_path.lstrip('/'), safe='/')
            full_source_path = os.path.join(workbooks_dir, encoded_source_path)

            # Check if the path exists and is a directory
            if os.path.exists(full_source_path) and os.path.isdir(full_source_path):
                destination_project = get_project_ids(
                    destination_path, cloud_server)

                for root, dirs, files in os.walk(full_source_path):
                    if root != full_source_path:
                        continue
                    for file in files:
                        file_path = os.path.join(root, file)
                        if os.path.isfile(file_path):
                            # Check if there is no extension and add ".twbx" for publication
                            if not os.path.splitext(file_path)[1]:
                                new_file_path = file_path + ".twbx"
                                os.rename(file_path, new_file_path)
                                file_path = new_file_path

                            # Sanitize and rename the file if necessary
                            file_name = os.path.basename(file)
                            decoded_file_name = urllib.parse.unquote(file_name)
                            sanitized_file_name = sanitize_filename(
                                decoded_file_name)
                            decoded_file_path = os.path.join(
                                root, sanitized_file_name)

                            if file_path != decoded_file_path:
                                os.rename(file_path, decoded_file_path)

                            converted_file_path = convert_path_format(
                                file_path)
                            for workbook_detail in workbook_details:
                                workbook_detail['filepath'] = convert_path_format(
                                    workbook_detail['filepath'])
                                if workbook_detail['filepath'] == converted_file_path and source_path == workbook_detail['datasource_projectpath']:
                                    datasource_projectpath = destination_path.split(
                                        "/")
                                    datasource_project_id = get_project_id(
                                        TSC.Pager(cloud_server.projects), datasource_projectpath, cloud_server)
                                    try:
                                        workbooks = Workbook(decoded_file_path)
                                        # print("Total Datasources in workbook : ", len(workbooks.datasources))
                                        for wbdatasource in workbooks.datasources:
                                            for wbconnection in wbdatasource.connections:
                                                # Update the 'destination_content_url' column for matching rows
                                                if not df.empty:
                                                    # Retrieve destination_content_url where Src_Content_Url equals dbname_value
                                                    matched_values = df.loc[df['Server_Src_Content_Url'] == str(
                                                        wbconnection.dbname), 'Cloud_Src_Content_Url'].values
                                                    # Check if there is a match and handle NaN values
                                                    matched_value = None  # Initialize the variable to store the matched value

                                                    if len(matched_values) > 0:
                                                        temp_value = matched_values[0]
                                                        # Check if the value is not NaN
                                                        if pd.notna(temp_value):
                                                            matched_value = temp_value

                                                    if matched_value is not None:
                                                        wbconnection.dbname = matched_value
                                                wbconnection.connection_type = workbook_detail['connection_type']
                                                server_address = workbook_detail['server_address']
                                                # wbconnection.server = workbook_detail['connection_type']
                                        workbooks.save()
                                        break
                                    except Exception as e:
                                        logger_error(f"Error: {e}")
                            if not os.path.splitext(decoded_file_path)[1]:
                                new_file_path = decoded_file_path + ".twbx"
                                os.rename(decoded_file_path, new_file_path)
                                decoded_file_path = new_file_path

                            workbook_name = ''
                            if decoded_file_name.endswith('.twbx'):
                                workbook_name = decoded_file_name[:-5]
                            elif decoded_file_name.endswith('.twb'):
                                workbook_name = decoded_file_name[:-4]
                            else:
                                workbook_name = decoded_file_name
                            workbook_name = workbook_name.replace(
                                '_fwd_SLASH_', '/')

                            if destination_project:
                                try:
                                    logger_info(
                                        f"Publishing {workbook_name} workbook to {destination_path} project")
                                    workbook_item = WorkbookItem(
                                        destination_project, name=workbook_name, show_tabs=True)
                                    published_workbook = cloud_server.workbooks.publish(
                                        workbook_item,
                                        decoded_file_path,
                                        'Overwrite',
                                        skip_connection_check="True",
                                    )
                                    # Log successful publication
                                    logger_info(
                                        f"\t\tWorkbook published successfully: {workbook_name}")
                                    workbooksname.append(workbook_name)
                                    log.append(
                                        f"Workbook migrated successfully")
                                    migrate.append('Migrated Workbook')
                                    total_specific_published += 1  # Increment counter
                                    set_description = ''
                                    for description in descriptions:
                                        if file_path.endswith('.twbx'):
                                            rename_file_path = file_path[:-5]
                                        elif file_path.endswith('.twb'):
                                            rename_file_path = file_path[:-4]
                                        description['published_workbook_id'] = ''
                                        rename_file_path = rename_file_path.replace(
                                            '/', '\\')
                                        if description['filepath'] == rename_file_path:
                                            description['published_workbook_id'] = published_workbook.id
                                            set_description = description['description']
                                            break
                                    description_status.append("")
                                    if set_description is not None:
                                        url = f"{cloud_server.baseurl}/sites/{cloud_server.site_id}/workbooks/{published_workbook.id}"
                                        headers = {
                                            "Content-Type": "application/xml",
                                            "X-Tableau-Auth": cloud_server.auth_token
                                        }
                                        # Assuming workbook_name and set_description are defined variables
                                        ts_request = ET.Element('tsRequest')

                                        workbook_elem = ET.SubElement(ts_request, 'workbook', {
                                            'description': set_description
                                        })

                                        # Convert the ElementTree object to a string representation of XML
                                        payload = ET.tostring(
                                            ts_request, encoding='utf-8', method='xml')
                                        try:
                                            response = requests.put(
                                                url, headers=headers, data=payload, verify=verify_ssl)
                                            logger_info(
                                                "\t\tDescription added successfully")
                                            description_status[-1] = "Description added successfully"
                                        except TSC.ServerResponseError as server_err:
                                            logger_error(
                                                f"Error publishing {workbook_name}: {server_err}")
                                            # Update the last status
                                            description_status[-1] = f"Error publishing {workbook_name}: {server_err}"
                                    publish_custom_view(
                                        cloud_server, descriptions, user_list, custom_temp_dir, published_workbook.id, cloud_users)
                                except TSC.ServerResponseError as server_err:
                                    logger_error(
                                        f"Error publishing {workbook_name}: {server_err}")
                                    workbooksname.append(workbook_name)
                                    migrate.append(
                                        'Failed to Migrate Workbook')
                                    log.append(
                                        f"Failed to Migrate Workbook: {server_err}")
                                    # Add failed workbook name to the list
                                    failed_workbooks.append(workbook_name)
                                    total_specific_failed += 1  # Increment counter
                            else:
                                logger_error(
                                    f"Project '{destination_path}' does not exist on the cloud server.")
                                logger_error(
                                    f"{workbook_name} workbook not published")
                                workbooksname.append(workbook_name)
                                log.append(
                                    f"Project '{destination_path}' does not exist on the cloud server.")
                                migrate.append('Failed to Migrate Workbook')
                                # Add workbook name to the list
                                failed_workbooks.append(workbook_name)
                                total_specific_failed += 1  # Increment counter
            else:
                logger_error(
                    f"Source '{source_path}' project dosen't have any workbook.")

        log_data = pd.DataFrame({
            'Workbook Name': workbooksname,
            'Log': log,
            'Status': migrate,
            'Description Status': description_status
        })

        file_name = 'migration_logs.xlsx'
        sheet_name = 'workbook'
        current_directory = os.getcwd()
        file_path = os.path.join(current_directory, file_name)

        try:
            workbook = openpyxl.load_workbook(file_path)
        except FileNotFoundError:
            workbook = openpyxl.Workbook()
            default_sheet = workbook["Sheet"]
            workbook.remove(default_sheet)

        if sheet_name not in workbook.sheetnames:
            workbook.create_sheet(sheet_name)

        sheet = workbook[sheet_name]
        sheet.delete_rows(1, sheet.max_row)

        sheet.append(log_data.columns.tolist())
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        for row in log_data.itertuples(index=False, name=None):
            sheet.append(row)
        workbook.save(file_path)

        logger_info(f"Migration log file generated: {file_path}")
        logger_info("-----End of publishing specific workbooks-----")

        if failed_workbooks:
            logger_error("Workbooks failed to publish:")
            for workbook in failed_workbooks:
                logger_error(f"- {workbook}")

    except Exception as migrate_error:
        logger_error(
            f"An error occurred during workbook migration: {migrate_error}")


def main():
    """
    Main executor function
    :return:
    """
    logger_info(f"::::Starting the migration workbooks ::::")
    start_time = system_time.time()

    parser = argparse.ArgumentParser(description="Read and parse an XML file.")
    parser.add_argument("--config", type=str, default=None,
                        help="Path to the XML file.")
    parser.add_argument("--project", type=str, default=None,
                        help="Path to the XLSX file.")
    parser.add_argument("--user", type=str, default=None,
                        help="Path to the .xlsx file containing user data.")
    args = parser.parse_args()
    userpath = args.config
    csv_userpath = args.project
    excel_user = args.user

    if userpath:
        try:
            path = os.path.join(userpath, 'config.xml')
            config = ReaderFactory(path).reader().to_json()
        except FileNotFoundError as e:
            logger_error(
                f"The specified config file does not exist in : {path}")
            exit(1)
    else:
        user_path = os.path.join(os.user_path.expanduser("~"), 'config.xml')
        config = ReaderFactory(user_path).reader().to_json()

    if csv_userpath:
        try:
            path = os.path.join(csv_userpath, 'project_path.xlsx')
            reader_factory = ReaderFactory(path)
            reader = reader_factory.reader()
            reader.read()
            csv_data = reader.to_json()
        except FileNotFoundError as e:
            logger_error(
                f"The specified config file does not exist in : {path}")
            exit(1)
    else:
        excel_path = os.path.join(
            os.excel_path.expanduser("~"), 'project_path.xlsx')
        csv_data = ReaderFactory(excel_path).reader().to_json()

    # Parse the JSON string into a list of dictionaries
    csv_data = json.loads(csv_data)

    if excel_user:
        try:
            path_u = os.path.join(excel_user, 'user_mapping.xlsx')
            user_file = pd.read_excel(path_u, sheet_name='Users')
        except FileNotFoundError as e:
            logger_error(
                f"The specified user mapping file does not exist in: {path}")
            exit(1)
    else:
        path = os.path.join(os.path.expanduser("~"), 'user_mapping.xlsx')
        user_file = ReaderFactory(path).reader().to_json()

    project_paths = [entry["Source Project_path"]
                     for entry in csv_data if entry.get("Select", "").lower() == "yes"]
    # # Download the Workbook from the server
    download_path, workbook_details, description, custom_temp_dir = download(
        config, project_paths)
    if download_path:
        print(f"Workbook downloaded to: {download_path}")
    else:
        print(f"Failed to download Workbook for project: {project_paths}")

    paths = [[entry["Source Project_path"], entry["Destination Project_path"]]
             for entry in csv_data if entry.get("Select", "").lower() == "yes"]

    cloud_server, credentials = login_to_tableau(
        config['credentials']['cloud'])
    with cloud_server.auth.sign_in(credentials):
        migrate_workbooks(cloud_server, download_path, paths, workbook_details,
                          description, user_file.values.tolist(), custom_temp_dir)
        end_time = system_time.time()
        total_time_seconds = end_time - start_time
        # Convert total_time_seconds to minutes and seconds
        minutes = int(total_time_seconds // 60)
        seconds = int(total_time_seconds % 60)
        # Log the total time
        logger_info(
            f"::::Execution time for migration workbooks is: {minutes} min {seconds} sec::::"
        )


if __name__ == "__main__":
    main()
