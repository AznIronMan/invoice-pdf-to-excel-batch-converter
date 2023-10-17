import datetime
import os

from datetime import datetime
from dotenv import load_dotenv
from typing import Optional, Union, Tuple


load_dotenv()


Detailed_Result = Tuple[
    bool,
    Optional[
        Union[str, tuple[float, float], tuple[int, int], tuple[str, str]]
    ],
]


def check_function(
    path: str, create_dir: bool = False, is_directory: bool = True
) -> Detailed_Result:
    try:
        if os.path.exists(path):
            return True, None
        else:
            if is_directory:
                if create_dir:
                    os.makedirs(path)
                    return True, None
                else:
                    raise Exception(f"{path} does not exist")
            else:
                raise Exception(f"{path} does not exist")
    except Exception as e:
        error_message = (
            "Failed to check "
            f"{'directory' if is_directory else 'file'}. Exception: {e}"
        )
        log(error_message, "WARNING")
        return False, error_message


def dir_check(dir_path: str, create_dir: bool = True) -> Detailed_Result:
    return check_function(dir_path, create_dir, is_directory=True)


def fix_datetime(
    input_time: datetime, milliseconds: bool = False
) -> Optional[str]:
    try:
        if isinstance(input_time, (int, float)):
            input_time = datetime.fromtimestamp(input_time / 1000)
        else:
            input_time = datetime.strptime(
                str(input_time), "%Y-%m-%d %H:%M:%S.%f"
            )
        if milliseconds:
            return input_time.strftime("%Y-%m-%d %H:%M:%S.%f")
        else:
            return input_time.strftime("%Y-%m-%d %H:%M:%S")
    except (ValueError, TypeError) as e:
        error_message = f"Failed to fix datetime. Exception: {e}"
        zlog(error_message, "ERROR")
        return None


def format_log_date(date_to_format) -> str:
    return date_to_format.replace("-", "")


def log(error_message: str, level: Optional[str] = "CRITICAL", success: bool = False) -> None:
    log_dir = f'../{os.getenv("LOG_DIR", "logs")}'
    current_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.join(current_dir, log_dir)
    dir_check(log_dir)
    log_file_stamp = format_log_date(today())
    log_file = os.path.join(log_dir, f"{log_file_stamp}.log")

    if dir_check(log_dir):
        is_new_file = not os.path.exists(log_file)

        with open(log_file, "a") as f:
            if is_new_file:
                f.write(f"[{now()}] New Log File Started.\n")

            if success:
                f.write(f"[{now()}] [{level}] Success: {error_message}\n")
            else:
                f.write(f"[{now()}] [{level}] Error: {error_message}\n")


def now() -> Optional[str]:
    return fix_datetime(datetime.utcnow(), milliseconds=True)


def today() -> Optional[str]:
    formatted_date = fix_datetime(datetime.utcnow())
    if formatted_date is None:
        error_message = "Failed to retrieve current date. Return is None."
        zlog(error_message, "ERROR")
        return None
    else:
        return formatted_date.split(" ")[0]


def zlog(
    exception: Union[Exception, str],
    level: str,
    success: bool = False,
    console: bool = False,
) -> None:
    exception_message = str(exception)
    force_debug = bool(os.getenv("FORCE_DEBUG", False))
    message = exception_message
    if isinstance(exception, Exception):
        message += f" of type {type(exception).__name__}"
    if console or force_debug:
        print(message)
    log(message, level, success)
