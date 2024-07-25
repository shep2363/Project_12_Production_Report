import os
import xml.etree.ElementTree as ET

def search_text_in_xml(directory, search_text):
    # Check if we can access the directory and it's not empty
    if not os.listdir(directory):
        print("The directory is empty or cannot be accessed.")
        return

    print(f"Searching for '{search_text}' in xml3Dm files within the directory: {directory}")

    # Counter for the number of xml3Dm files
    file_count = 0

    # Navigate through each file in the directory
    for filename in os.listdir(directory):
        # Check if the file is an xml3Dm file
        if filename.endswith('.xml3Dm'):
            file_count += 1
            print(f"Processing file: {filename}")

            # Construct the file path
            file_path = os.path.join(directory, filename)
            
            try:
                # Parse the XML
                tree = ET.parse(file_path)
                root = tree.getroot()

                # Flag to know if we found the text
                found_text = False

                # Recursively search for the text in the XML elements
                for elem in root.iter():
                    # Check if the element's text matches the search text
                    if elem.text is not None and search_text in elem.text:
                        print(f"Match found in file: {filename}")
                        found_text = True
                        # Break after finding the text, remove this if you want to search the whole file
                        break

                if not found_text:
                    print(f"No match found in file: {filename}")
            
            except ET.ParseError as e:
                print(f"Could not parse {filename}: {e}")
            except Exception as e:
                print(f"An error occurred while processing {filename}: {e}")

    if file_count == 0:
        print("No xml3Dm files found in the directory.")

def main():
    # Get the directory containing xml3Dm files from the user
    directory = input("Enter the directory path where xml3Dm files are stored: ").strip()
    
    # Validate that the directory exists
    if not os.path.exists(directory):
        print("The specified directory does not exist.")
        return
    
    # Get the text to search for from the user
    search_text = input("Enter the text you want to search for: ").strip()
    
    # Search for the text in the XML files
    search_text_in_xml(directory, search_text)

# Entry point of the script
if __name__ == "__main__":
    main()
