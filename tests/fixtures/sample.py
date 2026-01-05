# Sample Python file for testing

def simple_function():
    print("Hello")


def large_function(x, y, z):
    a = 1
    b = 2
    c = 3
    if x > 0:
        print("x is positive")
    else:
        print("x is not positive")
    return a + b + c + x + y + z


async def async_function():
    await some_async_call()
    return "async result"


def typed_function(name: str) -> str:
    return f"Hello, {name}"


def function_with_complex_return() -> list[dict[str, int]]:
    return [{"a": 1}, {"b": 2}]


class MyClass:
    def method_in_class(self):
        print("method")

    def another_method(self, x: int) -> int:
        return x * 2

    async def async_method(self):
        await some_async_call()
        print("async method")


# Nested function test
def outer_function():
    print("outer")

    def inner_function():
        print("inner")

    inner_function()
