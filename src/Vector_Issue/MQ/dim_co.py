#/usr/bin/python

# Benoetigte Bibiotheken einbinden
import sys
import os.path
import getpass
import subprocess
import Vector_Issue.MQ.dm as dm

# Anzahl der Parameter pruefen
if len(sys.argv) < 2:
  print "syntax: dim_co <file> <crs>"
  exit(0)

# Pruefen, ob von STDIN gelesen werden soll
elif sys.argv[1] == '-':
  fd = sys.stdin

# Pruefen, ob die angegebene Datei existiert
elif not os.path.isfile(sys.argv[1]):
  print "file %s does not exist" % sys.argv[1]
  exit(0)

# Datei oeffnen
else:
  fd = open(sys.argv[1], 'r')

# Alle angegebenen CRs ermitteln
crs = []
for cr in sys.argv[2:]:
  crs += [cr]

password = getpass.getpass("Password: ")
dm = dm.client('M306106', password)

# Alle Zeilen der Datei auslesen
fl = []
for line in fd:
  fl += [line.split()[-1]]
dm.co(fl, crs)

# Datei wieder schliessen
fd.close()
