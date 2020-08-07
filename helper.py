class CommonHelper:
    def __init__(self):
        pass

    @staticmethod
    def read_css(style):
        with open(style, 'r') as f:
            return f.read()
