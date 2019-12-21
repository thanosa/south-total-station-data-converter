# South Total Station NTS-350/350R Series data converter

## Input file format specs
- File extension is .dat
- First column starts in character 1
- Second column starts in character 9
- Fields in the second column are separated by commas

## Input file example
    JOB     SAMPLE
    INST    NTS-350 Ver. 2008.07.08
    UNITS   M,G
    STN     S2,1.636, 
    XYZ     212704.739,815754.136,556.168
    BS      S1,1.800, 
    SD      0.0000,96.9854,58.876
    BS      S1,1.800, 
    SD      399.9996,96.9854,58.908
    SS      1,0.000,2P
    SD      389.4092,96.6598,43.175
    SS      2,0.000,2P
    SD      385.1540,96.6601,31.506
    SS      3,0.000,5P

## Output file example
	ST,S2,1.636,S1
	SS,S1,0.0000,96.9854,58.876,1.800,
	SS,S1,399.9996,96.9854,58.908,1.800,
	SS,1,389.4092,96.6598,43.175,0.000,2P
	SS,2,385.1540,96.6601,31.506,0.000,2P
	SS,3,379.0913,97.2524,21.847,0.000,5P


## Pseudocode
The algorithm to convert the input dat file to a regular csv is the following:

1.  Read input file line by line
    1.  Strip line
    2.  Store line into a list element
2.  Initialize the output list
3.  Set read line index to 0
4.  While True
    1.  If the read line index is greater than the last
        1.  Break
    2.  Read current line
    3.  Read the key from the 1st column of the line
    4.  Read the value list from the 2nd column of the line
    5.  If the key is empty
        1.  Exit
    6.  Else-if the is in ignore list {“JOB”, ”INST”, “UNITS”, “XYZ”, “HV”}
        1.  Increase the real line index by 1
        2.  Continue
    7.  Else-if the key is “STN”
        1.  Read the first field from the value list as A1
        2.  Read the second field from the value list as A2
        3.  While True
            1.  Increase read line index by 1
            2.  Read current line
            3.  Read the key from the line
            4.  If the key is “BS”
                1.  Read the first field from the value list as B1
                2.  Write “ST,\<A1\>,\<A2\>,\<B1\>”
                3.  Break
    8.  Else-if key is “BS” or “SS”
        1.  Read the first field from the value list as K1
        2.  Read the second field from the value list as K2
        3.  Read the Third field from the value list as K3
            1.  While True
                4.  Increase the line index by 1
                5.  Read the current line
                6.  Read the key from the line
                7.  Read the value list from the line
                    1.  If key is “SD”
                        1.  Read the first field from the value list as L1
                        2.  Read the second field from the value list as L2
                        3.  Read the Third field from the value list as L3
                        4.  Write “SS,\<K1\>,\<K2\>,\<K3\>”
                        5.  Increase read line index by 1
                        6.  Break
    9.  Else
        1.  Increase search line index by 1
5.  Write the output list to the output file
