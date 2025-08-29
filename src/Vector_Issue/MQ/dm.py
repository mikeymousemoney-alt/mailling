import getpass
import subprocess

class client:
  """ Dimensions client class """
  def __init__(self, prd=None, pwd=None):
    # initialize internal variables
    self.prg = 'dmcli'
    self.usr = getpass.getuser()
    self.pwd = pwd
    self.host = 'dimensions'
    self.db = 'mq1'
    self.dsn = 'ae6'
    self.prd = prd

  # method for checking out files
  def co(self, fl, crs=[]):
    cmd = 'EI %s:; /overwrite /nometadata' % self.getprd()
    if not isinstance(fl, list):
      fl = [fl]
    if not isinstance(crs, list):
      crs = [crs]
    if len(crs) > 0:
      cmd += ' /change_doc_ids=%s' % ",".join(crs)
    cmd += ' /file='
    tmp = ''
    for fn in fl:
      tmp += cmd + '"' + fn + '"\n'
    self.run(tmp)

  # method for relating files to requests
  def rel(self, fl, ids):
    cmd = 'RICD %s:; /in_response /file="FN" /change=("ID")' % self.getprd()
    if not isinstance(fl, list):
      fl = [fl]
    if not isinstance(ids, list):
      ids = [ids]
    tmp = ''
    for fn in fl:
      tmp += cmd.replace('FN', fn).replace('ID', '", "'.join(ids)) + '\n'
    self.run(tmp)

  # internal method for running commands
  def run(self, cmd):
    # build string for execution
    str = [self.prg, '-user', self.usr, '-pass', self.getpwd(), '-host',
        self.host, '-dbname', self.db, '-dsn', self.dsn]
    prog = subprocess.Popen(str, stdin=subprocess.PIPE)
    prog.communicate(cmd + 'exit\n')
    prog.wait()

  # internal method for retrieving product id
  def getprd(self):
    # check current product
    if not self.prd:
      prd = raw_input('enter product ID: ')
      # check user input
      if len(prd):
        self.prd = prd
    return self.prd

  # internal method for retrieving password
  def getpwd(self):
    # check password
    if not self.pwd:
      pwd = getpass.getpass('enter password for user "%s": ' % self.usr)
      # check user input
      if len(pwd):
        self.pwd = pwd
    return self.pwd
