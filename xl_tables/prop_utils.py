
__all__ = ['ProxyProperty', 'ProxyMethod', 'HashDict']


class ProxyProperty(property):
    def __init__(self, attr_chain=None, fget=None, fset=None, fdel=None, doc=None):
        # Check if fget, fset, fdel given
        if callable(attr_chain):
            # Swap arguments to assume normal property was given
            doc, fdel, fset, fget = fdel, fset, fget, attr_chain
        else:
            # Create fget and fset from given string attribute chain
            attrs = str(attr_chain).split('.')
            if fget is None and isinstance(attr_chain, str):
                def fget(self):
                    obj = self
                    for attr in attrs:
                        obj = getattr(obj, attr)
                    return obj

            if fset is None and isinstance(attr_chain, str):
                def fset(self, value):
                    obj = self
                    for attr in attrs[:-1]:
                        obj = getattr(obj, attr)
                    setattr(obj, attrs[-1], value)

        super().__init__(fget, fset, fdel, doc)


class ProxyMethod(property):
    def __init__(self, attr_chain=None, fget=None, fset=None, fdel=None, doc=None, setter_error='Cannot set method '):
        # Check if fget, fset, fdel given
        if callable(attr_chain):
            # Swap arguments to assume normal property was given
            doc, fdel, fset, fget = fdel, fset, fget, attr_chain
        else:
            # Create fget and fset from given string attribute chain
            attrs = str(attr_chain).split('.')
            if fget is None and isinstance(attr_chain, str):
                def fget(self):
                    obj = self
                    for attr in attrs:
                        obj = getattr(obj, attr)
                    return obj

            if fset is None and isinstance(attr_chain, str):
                def fset(self, value):
                    if setter_error is not None:
                        raise AttributeError('{}{}.'.format(setter_error, attr_chain))
                    else:
                        # If setter_error is None do not raise error and call the setter function
                        func = self
                        for attr in attrs:
                            func = getattr(func, attr, None)
                        if callable(func):
                            func(value)
                        else:
                            raise AttributeError('{}{}.'.format('Invalid method! ', attr_chain))

        super().__init__(fget, fset, fdel, doc)


class HashDict(dict):
    # https://stackoverflow.com/questions/1151658/python-hashable-dicts
    def __hash__(self):
        return hash(tuple(sorted(self.items())))


class ItemStorage(list):
    def __init__(self, compare_func=None, *args, **kwargs):
        """Create item storage that can retrieve values based on custom settings.

        Args:
            compare_func (callable/function)[None]: Takes in two value compares them and returns True if they are the same.
        """
        self.compare_func = compare_func
        super().__init__(*args, **kwargs)

    def __getitem__(self, item):
        if isinstance(item, int):
            return super().__getitem__(item)

        call_compare = callable(self.compare_func)
        for val in self:
            if (not call_compare and val == item) or (call_compare and self.compare_func(val, item)):
                return val

        raise KeyError('Item not found!')
