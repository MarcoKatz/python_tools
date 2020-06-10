# -*- coding: utf-8 -*-
"""
Created on Fri Jun  5 22:57:32 2020

@author: Marco Katz
"""

def print_in_tab(title,headers,values):

    """Prints n number of headers and values in tabular format"""
    
    """
    Args:
        title(str) : if empty then title is not printed
        headers(list of str)
        values(list of str)
    """

    # Ensure all headers and values in the list are strings, and if not, convert
    for i, header in enumerate(headers):      
        if type(header) != "str": headers[i] = str(header) 
    for i, value in enumerate(values):      
        if type(value) != "str": values[i] = str(value)         

    # Check number of headers and number of values in input lists
    # If unequal, fill in gaps with "missing"
    num_headers=len(headers)
    num_values=len(values)
    if num_headers > num_values:
        for i in range(num_headers-num_values):
            values.append("Missing")
    elif num_headers < num_values:
        for i in range(num_values-num_headers):
            headers.append("Missing")
    
    # This is final number of columns
    column_count = max(num_headers,num_values)
    
    # Count 1 vertical bar before each entry, increase length of box by as much (minus 1 !)
    len_box = column_count-1
    
    # If there is a title, then print it
    if len(title) !=0:
        print(' '*5,title)
    
    # Initiate list of table column lengths, and contstruction of the print positioning string
    l_items=[]
    pos_var=[]
    pos_string = "{:<5}|"

    
    for i in range(0,column_count):
        # Construct a list with the width of each column (give each column 2 extra spaces)
        l_items.append(max(len(headers[i]),len(values[i]))+2)
        # Construct formatting string and grid elements
        pos_var.append("{:<"+str(l_items[i])+"}")
        pos_string += pos_var[i]+"|"
        len_box += l_items[i]
    
    # If box is too long to fit on a screen, just do a regular print
    if len_box > 85:
        print(' '*5,headers)
        print(' '*5,values)
    else:
                                        # Assume an indent of 5
            print(' '*5,'-'*len_box)                # Print the top of the box
            print(pos_string.format(" ",*headers))  # Print the line with the headers 
            print(' '*5,'-'*len_box)                # Print the middle of the box
            print(pos_string.format(" ",*values))   # Print the line with the values 
            print(' '*5,'-'*len_box)                # Print the bottom of the box

def main():

    headers = ['City','Month','Day','Hour','Minute']
    values = ['Washington DC','February','Wednesday','17','12']
    print_in_tab("Your selection",headers,values)


if __name__ == "__main__":
	main()
