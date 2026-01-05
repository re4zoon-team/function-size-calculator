# Sample Python test file for testing test file exclusion
# This file should be excluded from scanning

def test_simple_function():
    assert simple_function() is None


def test_large_function():
    result = large_function(1, 2, 3)
    assert result == 12


class TestMyClass:
    def test_method_in_class(self):
        obj = MyClass()
        obj.method_in_class()

    def test_another_method(self):
        obj = MyClass()
        assert obj.another_method(5) == 10
