title = "ERROR WARNING LIST:\n"
test_list = ['test1', 'test2', 'test3', 'test4']
with open('C:\\Users\\david\\desktop\\test_file_storage\\test1.txt', 'w') as f:
    for test in test_list:
        test_var = test
        test_var += '\n'
        f.write(test_var)
