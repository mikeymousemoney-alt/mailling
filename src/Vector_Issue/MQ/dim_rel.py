#/usr/bin/python

# Benoetigte Bibiotheken einbinden
import sys
import os.path
import getpass
import subprocess
import Vector_Issue.MQ.dm as dm

# Anzahl der Parameter pruefen
if len(sys.argv) < 3:
  print "syntax: dim_rel <file> <requests>"
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

dm = dm.client('M306106', 'sebastian09')
ids = sys.argv[2:]

# Alle Zeilen der Datei auslesen
fl = []
for line in fd:
  fl += [line.split()[-1]]
dm.rel(fl, ids)

# Datei wieder schliessen
fd.close()
