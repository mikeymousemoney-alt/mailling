test_mode = 0
def set_test_mode(mode):
    global test_mode
    test_mode = mode
# prints for debugging if testmode is enabled
def test_log(*args):
    if test_mode == 1:
        print(*args)

def get_test_mode():
    return test_mode