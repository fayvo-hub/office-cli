"""统一的业务错误类型。"""


class CliError(Exception):
    """用户输入 / 文件 / 业务规则错误。

    消息直接面向 AI 展示,要求:中文、可操作、给出修正建议。
    """

    def __init__(self, code: str, message: str, *, exit_code: int = 3):
        super().__init__(message)
        self.code = code
        self.message = message
        self.exit_code = exit_code
