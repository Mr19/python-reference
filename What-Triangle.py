#####
# Written by Gaddy
#####

#!/usr/bin/env python
#

import optparse
import sys

def formula():
    if options.value[0]*options.value[0] + options.value[1]*options.value[1] == options.value[2]*options.value[2]:
        print("Right Angled Triangle")
    elif options.value[0]*options.value[0] + options.value[1]*options.value[1] > options.value[2]*options.value[2]:
        print("Isocles Triangle")
    else:
        print ("Equilateral Triangle")

def _build_cli_parser():
    usage = "usage: %prog [options]"
    desc = "Find the triangle"

    parser = optparse.OptionParser(usage=usage, description=desc)

    parser.add_option("-i", "--input", action="store", type="int", dest="value", nargs=3, help="Input 3 values")

    return parser

if __name__ == "__main__":

    parser = _build_cli_parser()
    options, args = parser.parse_args(sys.argv)
    formula()
    sys.exit(0)
