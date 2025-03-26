# db_connection.py
import psycopg2

from tableauserverclient import PersonalAccessTokenAuth,Server,ServerResponseError,NotSignedInError
from tableauserverclient import Pager
from logger.logger import *
import requests
from pypac import pac_context_for_url


def get_db_connection(database_credentials):
    try:
        connection = psycopg2.connect(
            dbname=database_credentials['@dbname'],
            user=database_credentials['@user'],
            password=database_credentials['@password'],
            host=database_credentials['@host'],
            port=database_credentials['@port']
        )
        return connection
    except (Exception, psycopg2.Error) as error:
        logger_error(f"Error while connecting to PostgreSQL: {error}")
        exit()


def login_to_tableau(config):
    # Extract the PAC URL from the configuration file
    pac_url = config.get("@pac_url", "")
    logger_info(f"PAC URL: {pac_url}")

    # Use pac_context_for_url only if pac_url is not empty or None
    if pac_url:
        with pac_context_for_url(pac_url):
            return _login_with_context(config)
    else:
        return _login_with_context(config)

def _login_with_context(config):
    try:
        credentials = PersonalAccessTokenAuth(
            token_name=config["@pat_name"],
            personal_access_token=config["@pat"],
            site_id=config["@site"]
        )
        server = Server(config["@path"], use_server_version=True)
        server.version="3.21"

        # SSL verification handling
        if config["@ssl"] in ["False", "false", "No", "no"]:
            server.add_http_options({'verify': False})
        else:
            server.add_http_options({'verify': config['@ssl_filepath']})

        return server, credentials

    except ServerResponseError as error:
        logger_error(f"Failed to login: {error}")
        exit(0)
    except NotSignedInError as error:
        logger_error(f"Sign-in error: {error}")
        logger_error("The personal access token you provided is invalid.")
        exit(0)
    except (requests.ConnectionError, requests.Timeout) as error:
        logger_error(f"Internet connection error: {error}")
        exit(0)
    except OSError as error:
        logger_error(f"Error occurred while reading file: {error}")
        exit(0)
    except Exception as error:
        logger_error(f"An unexpected error occurred: {error}")
        exit(0)


def get_all_items(server, item):
    items = {
        'project': server.projects,
        'workbook': server.workbooks,
        'datasource': server.datasources,
        'view': server.views,
        'flow': server.flows,
        'user': server.users,
        'group': server.groups
    }
    if item in items:
        return list(Pager(items[item]))

def get_project_ids(path, server):
    all_projects = get_all_items(server, 'project')
    path = path[1:]
    if '/' in path:
        path_parts = path.split('/')
        return get_child_project_id(all_projects, path_parts, server)
    else:
        for pro in all_projects:
            if pro.name == path and pro.parent_id is None:
                return pro.id

def get_child_project_id(projects, project_path, server):
    if not project_path:
        return None

    for project in projects:
        # Check if the current project name matches the first part of the project path
        if project.name == project_path[0]:
            # It's the last part of the path, return the project ID
            if len(project_path) == 1:
                return project.id
            # Otherwise, recursively search for the child project ID
            else:
                child_projects = [p for p in Pager(server.projects) if p.parent_id == project.id]
                try:
                    child_project_id = get_child_project_id(child_projects, project_path[1:], server)
                    if child_project_id:
                        return child_project_id
                except (IndexError, AttributeError) as e:
                    logger_error(f"Error: {e}")
                    return None

    # If no matching project is found, return None
    return None

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

    return full_path if full_path else 'Top-level-project'
