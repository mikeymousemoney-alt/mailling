
import pytest
from shutil import copy

def make_backup():
    copy('src/../config.txt', 'src/../config_backup.txt')

def restore_backup():
    copy('src/../config_backup.txt', 'src/../config.txt')


class TestMain:

    # Successfully opens 'src/Vector_Issue/config.txt' and processes its contents
    def test_successful_open_src_config(self):
        make_backup()
        import os
        import json
        from src.Vector_Issue.make_config import main
    
        # Setup
        os.makedirs('src/Vector_Issue', exist_ok=True)
        with open('src/../config.txt', 'w') as f:
            f.write('key1:value1\nkey2:value2\n')
    
        # Execute
        main()
    
        # Verify
        with open('src/Vector_Issue/config.json', 'r') as f:
            data = json.load(f)
            assert data == {'key1': 'value1', 'key2': 'value2'}
    
        # Cleanup
        os.remove('src/../config.txt')
        restore_backup()
        main()

    # Correctly removes empty lines and comments from the configuration file with the recommended fix
    def test_remove_empty_lines_and_comments_fixed(self):
        make_backup()
        import os
        import json
        from src.Vector_Issue.make_config import main

        # Setup
        os.makedirs('src/Vector_Issue', exist_ok=True)
        with open('src/../config.txt', 'w') as f:
            f.write('key1:value1\nkey2:value2\n\n////comment1\nkey3:value3\n')

        # Execute
        main()

        # Verify
        with open('src/Vector_Issue/config.json', 'r') as f:
            data = json.load(f)
            assert data == {'key1': 'value1', 'key2': 'value2', 'key3': 'value3'}

        # Cleanup
        os.remove('src/../config.txt')
        restore_backup()
        main()

    # Correctly splits lines into key-value pairs and stores them in a dictionary
    def test_correctly_splits_lines_into_key_value_pairs(self):
        make_backup()
        from src.Vector_Issue.make_config import main
        import os
        import json

        # Setup
        os.makedirs('src/Vector_Issue', exist_ok=True)
        with open('src/../config.txt', 'w') as f:
            f.write('key1:value1\nkey2:value2\n')

        # Execute
        main()

        # Verify
        with open('src/Vector_Issue/config.json', 'r') as f:
            data = json.load(f)
            assert data == {'key1': 'value1', 'key2': 'value2'}

        # Cleanup
        os.remove('src/../config.txt')
        restore_backup()
        main()

    # Successfully writes the dictionary to 'src/Vector_Issue/config.json' if the first path is used
    def test_writes_to_src_config_json(self):
        make_backup()
        import os
        import json
        from src.Vector_Issue.make_config import main

        # Setup
        os.makedirs('src/Vector_Issue', exist_ok=True)
        with open('src/../config.txt', 'w') as f:
            f.write('key1:value1\nkey2:value2\n')

        # Execute
        main()

        # Verify
        with open('src/Vector_Issue/config.json', 'r') as f:
            data = json.load(f)
            assert data == {'key1': 'value1', 'key2': 'value2'}

        # Cleanup
        os.remove('src/../config.txt')
        restore_backup()
        main()

    # Dictionary keys or values contain leading or trailing whitespace
    def test_dictionary_keys_values_whitespace_fixed_fixed(self):
        make_backup()
        import os
        import json
        from src.Vector_Issue.make_config import main

        # Setup
        os.makedirs('src/Vector_Issue', exist_ok=True)
        with open('src/../config.txt', 'w') as f:
            f.write('key1 : value1\nkey2: value2\n')

        # Execute
        main()

        # Verify
        with open('src/Vector_Issue/config.json', 'r') as f:
            data = json.load(f)
            assert {k.strip(): v for k, v in data.items()} == {'key1': 'value1', 'key2': 'value2'}

        # Cleanup
        os.remove('src/../config.txt')
        restore_backup()
        main()