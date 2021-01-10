import csv

'''
Data Example
CSV Data
Component Name, BOM
Microchip 1, NI
Resistor 2, I
Capacitor 3, NI
'''

def get_list_of_dict_from_rpt(file_path):
    with open(file_path, newline='') as file:
        # Get list of field names
        for i in range(6):
            fields = file.readline()
            if i == 5:
                field_names = fields[:-1].split("#")
        
        reader = csv.DictReader(file, fieldnames=field_names, delimiter='#')

        rpt_list = []
        for row in reader:
            rpt_list.append(row)
    return rpt_list

def get_installed_components_from_dict_list(rpt_list):
    installed = []
    for item in rpt_list:
        if item['BOM'] == 'I' or item['BOM'] == 'DEBUG':
            installed.append(item)
    return installed

def get_not_installed_components_from_dict_list(rpt_list):
    not_installed = []
    for item in rpt_list:
        if item['BOM'] == 'NI':
            not_installed.append(item)
    return not_installed

def get_everything_else_from_dict_list(rpt_list):
    everything_else = []
    for item in rpt_list:
        if item['BOM'] != 'NI' and item['BOM'] != "I" and item['BOM'] != "DEBUG":
            everything_else.append(item)
    return everything_else

def write_to_csv(dict_list, output_path):
    keys = dict_list[0].keys()
    with open(output_path, 'w', newline='') as file:
        dict_writer = csv.DictWriter(file, keys)
        dict_writer.writeheader()
        dict_writer.writerows(dict_list)


if __name__ == "__main__":
    # Get data into list of dictionaries where each row is a dictionary
    csvpath = './CONAN_MB_EVT1_BOM_20210105.rpt'
    rpt_list = get_list_of_dict_from_rpt(csvpath)
    print(f"The length of the data is: {len(rpt_list)}")

    # Get CSV of installed BOM
    installed_list = get_installed_components_from_dict_list(rpt_list)
    print(f"The length of the installed components are: {len(installed_list)}")
    installed_output = './installed_BOM.csv'
    write_to_csv(installed_list, installed_output)

    # Get CSV for non-installed BOM
    not_installed_list = get_not_installed_components_from_dict_list(rpt_list)
    print(f"The length of the non-installed components are: {len(not_installed_list)}")
    not_installed_output = './not_installed_BOM.csv'
    write_to_csv(not_installed_list, not_installed_output)

    # Get CSV for non-installed BOM
    everything_else_list = get_everything_else_from_dict_list(rpt_list)
    print(f"The length of everything else components are: {len(everything_else_list)}")
    everything_else_output = './everything_else_BOM.csv'
    write_to_csv(rpt_list, everything_else_output)

    # There are BOM equal to DEBUG.
    # for item in rpt_list:
    #     if item['BOM'] != "I" and item['BOM'] != 'NI':
    #         print(item['BOM'])


