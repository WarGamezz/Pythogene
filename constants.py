#for whatever reason windows cmd requires clear screen command executed in order to ASCII escape codes (coloring) work correctly. Seems like WINDOWS OS specific issue.
import os

#**************************************** VISUALS / COLORS ****************************************# 
#define ANSI color escape codes
colors = {
        'blue': '\033[94m',
        'pink': '\033[95m',
        'green': '\033[92m',
        'red': '\033[31m',
        'yellow': '\033[33m',
        }

#console output colorizing function
def colorize(string, color):
    if not color in colors: return string
    return colors[color] + string + '\033[0m'

#**************************************** MENU ITEMS ****************************************# 

#all of the elements that are constantly listed on every page of the menu are listed below

def splash():
        #clear screen (saves time compared to positioning menu items manually under splash it's easier to re-draw everything)
        os.system('cls') 
        print(colorize('''

            ██████╗ ██╗   ██╗████████╗██╗  ██╗ ██████╗  ██████╗ ███████╗███╗   ██╗███████╗
            ██╔══██╗╚██╗ ██╔╝╚══██╔══╝██║  ██║██╔═══██╗██╔════╝ ██╔════╝████╗  ██║██╔════╝
            ██████╔╝ ╚████╔╝    ██║   ███████║██║   ██║██║  ███╗█████╗  ██╔██╗ ██║█████╗  
            ██╔═══╝   ╚██╔╝     ██║   ██╔══██║██║   ██║██║   ██║██╔══╝  ██║╚██╗██║██╔══╝            
            ██║        ██║      ██║   ██║  ██║╚██████╔╝╚██████╔╝███████╗██║ ╚████║███████╗           by YeltsinB
            ╚═╝        ╚═╝      ╚═╝   ╚═╝  ╚═╝ ╚═════╝  ╚═════╝ ╚══════╝╚═╝  ╚═══╝╚══════╝                  v0.1




        ''','green'))

def constant_menu():
        print('')
        print(colorize('(q) Quit ','yellow'))
        print('')
        print('--------------------------------')

#disclaimer showed at start
def disclaimer():
        print(colorize('''

             DISCLAIMER:

     This application was developed for training/etertainment purposes only. It's source is readily available
     and included in the files provided. Feel free to modify this program as you please. I hope that whoever you are
     you're going to find good use for it. Selenium and pyodbc modules are required to run this program correctly.

                                                                                                        Cheers!     



        ''', 'blue'))
        #wait for user input to proceed to the main program
        input(colorize('Press ENTER key to continue ...', 'yellow'))
