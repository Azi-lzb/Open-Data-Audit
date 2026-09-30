"""审核工具。"""

def __getattr__(name: str):
    if name == "__version__":
        from .app_identity import load_product_info

        return load_product_info()["version"]
    raise AttributeError(name)
