from importlib.metadata import Distribution

__pkg       = Distribution.from_name(__package__)
__version__ = __pkg.version

try:
    from .mbox2m365         import *
except:
    try:
        from .                  import mbox2m365
    except:
        import mbox2m365