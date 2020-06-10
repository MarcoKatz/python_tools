# -*- coding: utf-8 -*-
"""
Created on Fri Jun  5 07:48:24 2020

@author: Marco Katz
"""


def seek_choice(message,option_list,no_choice=False,all_choice=False,fold=True):
    
    """
    Collects 1 choice from a user, from within a list of options
    Uses a numerical identifier for each option
    The user can quit the selection if he/she wants to break the process

    Args:
    -   message (str): the question that is asked to the user
    -   option_list (list): the list (of strings) from which a selection must be made
    -   no_choice (boolean, default False): if True, activates 'None' as a valid choice
    -   all_choice (boolean default False): if True, activates 'All' as a valid choice
    -   fold (boolean, default True): if True, options are presented on different lines,
            If false, options are presented on the same line, separated by comma's
    Returns:
    -   choice (str): the chosen option, 'Quit', or 'None' and/or 'All' if applicable   
        
       
    
    """   
    
    print("\n" + message)
    
    # Initiate set up of the main message, the queue for the user's reply, and the list of valid choices
    fold_char = ("\n" if fold else ", ")
    option_message = ""
    choice_message = "("
    valid_choices = []
    
    # Build the messages and the list of valid choices 
    for i, option in enumerate(option_list):
        option_message += option + ": " + str(i+1) + fold_char
        choice_message += str(i+1) + ","
        valid_choices.append(str(i+1))
    
    # If None is an option, add this to messages and list of choices
    if no_choice:
        option_message += "None: N" + fold_char
        choice_message += "N,"
        valid_choices.append("N")
    
    # If All is an option, add this to messages and list of choices
    if all_choice:
        option_message += "All: A" + fold_char
        choice_message += "A,"
        valid_choices.append("A")
    
    # Quit is always an option
    option_message += "Quit: Q"
    choice_message += "Q"
    valid_choices.append("Q")
    
    # Print messages and collect input
    print(option_message)
    choice_message = "Your choice"+choice_message +"): "
    while True:
        try:
            choice_index = input(choice_message).title() # Expect a string in return
            if choice_index in valid_choices:
                if choice_index == "Q":
                    choice = 'Quit'
                elif choice_index == "A":
                    choice = 'All'
                elif choice_index == "N":
                    choice = 'None'
                else:
                    choice = option_list[int(choice_index)-1] # Pick up option corresponding to the input
                break
            print("Invalid choice: please re-try")
        except ValueError:
            print("Invalid input: please re-try")
        except KeyboardInterrupt:
            choice = 'Quit'
            break
    return choice


 

def main():
    
    
    city=seek_choice("Please select the city you want to see results for:",['Chicago','New York City','Washington DC'],False,True,True)
    filter=seek_choice("Do you want to filter by ... ",['Month','Week-Day'],True,True,False)
    month=seek_choice("Please select the month you wish to filter by",['January','February','March','April','May','June'],False,True,False)
    week_day=seek_choice("Please select the week-day you wish to filter by",['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'],False,False,False)

    print(city,filter,month,week_day)
    

if __name__ == "__main__":
	main()
