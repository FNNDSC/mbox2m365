import asyncio
import ssl
from threading import ExceptHookArgs
from aiosmtpd.controller import Controller, BaseController
from aiosmtpd.smtp import Envelope, Session, AuthResult, LoginPassword
import argparse
from pathlib import Path
import mailbox
import sys
import pudb
from email import message_from_bytes
from pydantic import BaseModel, field_validator, validator, Field


auth_db = {
    b"user1": b"password1",
    b"user2": b"password2",
    b"user3": b"password3",
}


def authenticator_func(server, session, envelope, mechanism, auth_data) -> AuthResult:
    # For this simple example, we'll ignore other parameters
    pudb.set_trace()
    assert isinstance(auth_data, LoginPassword)
    username = auth_data.login
    password = auth_data.password
    # If we're using a set containing tuples of (username, password),
    # we can simply use `auth_data in auth_set`.
    # Or you can get fancy and use a full-fledged database to perform
    # a query :-)
    if auth_db.get(username) == password:
        return AuthResult(success=True)
    else:
        return AuthResult(success=False, handled=False)


class MailServerArgs(BaseModel):
    mdir: Path = Field(..., title="Output mbox directory")
    mbox: Path = Field(..., title="Output mbox file")
    # mdir: Path = Path("/tmp/mail")
    # mbox: Path = Path("messages.mbox")
    tls: bool = False
    port: int = 22225

    @field_validator("mdir", "mbox")
    def validate_path(cls, v, field):
        if v == "mdir" and not v.is_dir():
            try:
                v.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                raise ValueError(f"{v} is not a valid directory path")
        return v


class Mbox365Handler:
    def __init__(self, mboxFile: Path):
        self.mboxFile: Path = mboxFile

    async def handle_DATA(
        self, server: Controller, session: Session, envelope: Envelope
    ) -> str:
        try:
            contents = envelope.content
            msg: bytes = message_from_bytes(contents)
            print(f"Message from {envelope.mail_from}")
            print(f"Message for {envelope.rcpt_tos}")
            print(f"Message data:\n{contents}")
            print("End of message")
            mbox: mailbox.mbox = mailbox.mbox(self.mboxFile, create=True)
            mbox.add(msg)
            mbox.flush()
            mbox.close()

            return "250 Message accepted for delivery"

        except Exception as e:
            pudb.set_trace()
            print(f"Unrecognized command or data received: {e}")
            return "500 Syntax error: command unrecognized"

    # async def handle_AUTH(
    #     self, server: Controller, session: Session, envelope: Envelope, args: bytes
    # ) -> AuthResult:
    #     pudb.set_trace()
    #     if args[0].decode() == "PLAIN":
    #         if args[1].decode() == "\x00ch137123@gmail.com\x00your_password":
    #             session.authenticated = True
    #             return AuthResult(success=True)
    #     return AuthResult(success=False)


def context_get() -> ssl.SSLContext:
    context: ssl.SSLContext = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.load_cert_chain("cert.pem", "key.pem")
    return context


def controller_get(args: MailServerArgs) -> Controller:
    controller: Controller
    if args.tls:
        controller = Controller(
            Mbox365Handler(args.mdir / args.mbox),
            port=args.port,
            authenticator=authenticator_func,
            ssl_context=context_get(),
        )
        print(f"TLS mailserver listening on {args.port}")
    else:
        controller = Controller(Mbox365Handler(args.mdir / args.mbox), port=args.port)
        print(f"No auth/no TLS mailserver listening on {args.port}")

    return controller


def env_check(args: MailServerArgs) -> bool:
    OK: bool = False
    mbox_path = Path(args.mdir) / Path(args.mbox)
    try:
        mbox_path.parent.mkdir(parents=True, exist_ok=True)
        OK = True
    except Exception as e:
        print(f"An error occured checking the env: {e}")

    return OK


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Mail sink server")
    parser.add_argument(
        "--mdir",
        required=False,
        help="Output mbox directory",
        default="/tmp/mail",
        type=str,
    )
    parser.add_argument(
        "--mbox",
        required=False,
        help="Output mbox file",
        default="messages.mbox",
        type=str,
    )
    parser.add_argument(
        "--tls",
        required=False,
        help="toggle TLS context",
        default=False,
        action="store_true",
    )
    parser.add_argument(
        "--port",
        required=False,
        help="Listener port",
        default=22225,
        type=int,
    )
    return parser.parse_args()


if __name__ == "__main__":
    options: argparse.Namespace = parse_arguments()
    args = MailServerArgs(**vars(options))

    if not env_check(args):
        sys.exit(1)

    print(f"All email will be sunk to {args.mdir / args.mbox}")

    controller: Controller = controller_get(args)
    controller.start()
    input("Press enter to stop")
    controller.stop()
