import pyrfc  # Library for connecting to SAP systems using RFC
import dotenv as dt  # Library for loading environment variables from .env file
import pprint as pp  # Library for pretty printing complex data structures
import os  # Library for interacting with the operating system, including environment variables

# Load environment variables from .env file
dt.load_dotenv()

def get_sap_data_rfc(import_parameters=None):
    """
    Connects to an SAP system using RFC, calls a specified RFC function, and returns the result.

    Args:
        import_parameters (dict, optional): A dictionary containing the input parameters for the RFC function. 
                                            Defaults to None.

    Returns:
        dict or None: The result returned by the RFC function, or None if an error occurs.
    """
    try:
        # Create a connection to the SAP system using environment variables for credentials
        config = pyrfc.Connection(**{
            'user': os.getenv('SAP_USER'),       # SAP username from environment variable
            'password': os.getenv('SAP_PASSWORD'), # SAP password from environment variable
            'ashost': os.getenv('SAP_ASHOST'),     # SAP application server hostname from environment variable
            'sysnr': os.getenv('SAP_SYSNR'),       # SAP system number from environment variable
            'client': os.getenv('SAP_CLIENT')      # SAP client number from environment variable
        })

        # Call the specified RFC function with the provided import parameters
        result = config.call("YOUR_RFC_FUNCTION_NAME", **import_parameters)  # Replace YOUR_RFC_FUNCTION_NAME with the actual function name

        # Close the SAP connection
        config.close()

        # Return the result from the RFC function
        return result

    except pyrfc.ABAPApplicationError as e: # Catch specific RFC errors (errors raised within the SAP system)
        print(f"Error calling RFC: {e}") # Print the error message
        return None # Return None to indicate an error


# Example usage:
import_parameters = {"MATERIAL": "your_material", "PLANT": "your_plant"}  # Example input parameters - replace with your actual parameters and names
                                                                     # Ensure the parameter names match the RFC function's signature.


sap_data = get_sap_data_rfc(import_parameters) # Call the function to retrieve data from SAP

# Pretty print the returned SAP data for easy readability
pp.pprint(sap_data) 
