from functools import wraps

def cache_on_attr(attr_name):
    cache_dict = {}

    def decorator(func):
        @wraps(func)
        def wrapper(obj, *args, **kwargs):
            key = getattr(obj, attr_name)
            if key in cache_dict:
                return cache_dict[key]
            result = func(obj, *args, **kwargs)
            cache_dict[key] = result
            return result
        return wrapper
    return decorator