# Adapted by Tim Collyer and Girithar from original perl script by Chris Crowley

import sys

index = {}

a = sys.argv[1].split('.')[0]

for file in sorted(sys.argv[2:]):
   if a == file.split('.')[0]:
       continue
   else:
       print 'All filenames need to begin with the same string, typically the number of the class e.g. "575.1 575.2 575.3"'
       sys.exit()

for file in sorted(sys.argv[1:]):
  with open(file, 'r') as fname:
     data = fname.readlines()
     for line in data:
         if line.rstrip().endswith(";"):
             line = line.rstrip()[:-1]
             val = line.split(';')
         else:
             val = line.rstrip().split(';')
         for each in val[1::]:
             if each in index.keys():
                 index[each].append(file + '.' + str(val[0]))
             else:
                 index[each] = [file + '.' + str(val[0])]

for key, val in sorted(index.iteritems()):
  print(key + ": " + ','.join(str(v) for v in val))
