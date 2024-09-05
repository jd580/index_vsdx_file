import re
import vsdx
import argparse
import sys

parser = argparse.ArgumentParser(description="Create an index in the current working directory of all IP addresses and networks in a vsdx file.")
parser.add_argument('--vsdx', action='store', metavar='NetworkDiagram.vsdx', help='Path to a copy of network diagram vsdx file.')

if len(sys.argv)==1:
    parser.print_help(sys.stderr)
    sys.exit(1)

args = parser.parse_args()
filename = args.vsdx

network_address_re = re.compile('(?:[0-9]{1,3}\.){3}[0-9]{1,3}/[0-9]{2}')
ip_address_re = re.compile('(?:[0-9]{1,3}\.){3}[0-9]{1,3}')

index_file = open('vsdx_index_file.csv','w+')
cols='page_name,type,match,raw_line\n'
index_file.write(cols)

# Open the visio file
with vsdx.VisioFile(filename) as vis:
    print("Found " + str(len(vis.pages)) + " pages in " + filename +"!")
    # Cycle through the pages.
    for page in vis.pages:
        print("Parsing page: " + page.name)
        # Search for shapes with IP addresses in the text portion of the shape.
        for shape in page._shapes:
            shapes_with_ip_addresses = shape.find_shapes_by_regex(ip_address_re)
            # Print the text.
            for lines in shapes_with_ip_addresses:
                # Search through the individual lines for network addresses or IP addresses.
                for ln in lines.text.splitlines():
                    # Save them to the index file.
                    if network_address_re.search(ln):
                        # Add entry to CSV of the form: page_name, type, match, raw_match
                        index_file.write(page.name + ",network," + network_address_re.search(ln).group() + "," + ln.replace(',', ';') + "\n")
                    elif ip_address_re.search(ln):
                        index_file.write(page.name + ",ipaddress," + ip_address_re.search(ln).group() + "/32," + ln.replace(',', ';') + "\n")

# Close the csv file.
index_file.close()
