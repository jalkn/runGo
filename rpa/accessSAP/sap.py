import dotenv as dt
import pprint as pp
import requests as rq
import os
import json  # For handling JSON data
import urllib.parse # For URL encoding

# Load environment variables from .env file
dt.load_dotenv()

def get_sap_data(query_parameters=None):  # Function to retrieve data from SAP API, accepts optional query parameters
    try:
        # 1. Construct the SAP API URL
        base_url = os.getenv("SAP_API_URL")  # Retrieve the base URL of the SAP API from environment variables
        
        if query_parameters:
            # Construct the query string from the provided parameters
            query_string = urllib.parse.urlencode(query_parameters) # Encode the query parameters to ensure URL validity
            url = f"{base_url}?{query_string}" # Append the query string to the base URL


        else:
            url = base_url # Use the base URL directly if no query parameters are provided

        # 2. Set up authentication 
        headers = {
            "Authorization": f"Bearer {os.getenv('SAP_API_TOKEN')}",  # Set the Authorization header with the API token from environment variables (Bearer token used here)
            "Content-Type": "application/json" # Set the Content-Type header, specifying that the request body is in JSON format
        }

        # 3. Make the API request
        response = rq.get(url, headers=headers)  # Make a GET request to the constructed URL with the specified headers
        response.raise_for_status()  # Raise an exception for bad status codes (4xx or 5xx), simplifying error handling

        # 4. Parse the response
        if response.headers['Content-Type'] == 'application/json': # Check the response's Content-Type header
            sap_data = response.json() # Parse the JSON response into a Python dictionary
        #elif response.headers['Content-Type'].startswith('text/xml'):  # Handle XML if needed - uncomment and adapt if the API returns XML
        #   sap_data = xml.etree.ElementTree.fromstring(response.content) # Parse the XML response - requires `import xml.etree.ElementTree` 
        else:
            raise ValueError("Unexpected content type") # Raise an exception if the Content-Type is neither JSON nor XML

        return sap_data # Return the parsed data

    except rq.exceptions.RequestException as e:  # Handle exceptions related to the API request (e.g., network issues)
        print(f"Error accessing SAP data: {e}") # Print the error message
        return None # Return None to indicate failure

    except json.JSONDecodeError as e: # Handle JSON decoding errors (if the response is not valid JSON)
        print(f"Error decoding JSON: {e}")  # Print the error message
        return None # Return None to indicate failure



if __name__ == "__main__": # Ensures that the code within this block only runs when the script is executed directly
    print('\n*** Get SAP Data ***\n')

    # Example: Get user input for query parameters
    material_number = input("Enter Material Number: ") # Prompt the user to enter a material number
    plant = input("Enter Plant: ") # Prompt the user to enter a plant

    query_parameters = { # Create a dictionary to store the query parameters
        "material": material_number, # Map the "material" parameter to the entered material number
        "plant": plant # Map the "plant" parameter to the entered plant
    }

    sap_data = get_sap_data(query_parameters) # Call the get_sap_data function with the query parameters

    if sap_data: # Check if the function returned data successfully
        pp.pprint(sap_data) # Pretty print the retrieved SAP data
