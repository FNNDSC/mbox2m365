import ssl
import asyncio
import email
import argparse
from typing import Optional
from pathlib import Path
from aiosmtpd.controller import Controller
from aiosmtpd.smtp import SMTP, Session, Envelope


class MailServer(SMTP):
    def __init__(self, mbox_file: Path):
        super().__init__(self.handle_DATA)
        self.mbox_file = mbox_file

    async def handle_DATA(
        self, server: Controller, session: Session, envelope: Envelope
    ):
        # Parse the email message
        msg = email.message_from_bytes(envelope.content)

        # Extract the email body
        body = msg.get_payload()

        # Save the email body to the mbox file
        with self.mbox_file.open("a") as mbox:
            mbox.write(f'From {envelope.mail_from} {msg["Date"]}\n')
            mbox.write(body.decode("utf-8"))
            mbox.write("\n\n")

        # Return a successful response
        return "250 OK"


async def run_server(host: str, port: int, mbox_file: Path):
    # Create the mail server instance
    mail_server = MailServer(mbox_file)

    # Create an SSL context for STARTTLS
    ssl_context = ssl.create_default_context(ssl.Purpose.CLIENT_AUTH)
    ssl_context.load_cert_chain(certfile="server.crt", keyfile="server.key")

    # Start the mail server with STARTTLS support
    server = Controller(mail_server, hostname=host, port=port, ssl_context=ssl_context)
    server.start()

    print(f"Mail server started on {host}:{port}")

    # Stop the server and wait for it to complete
    server.stop()
    await asyncio.wait([server.stopped])

    print("Mail server stopped")


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
        "--port",
        required=False,
        help="Listener port",
        default=8025,
        type=int,
    )
    return parser.parse_args()


if __name__ == "__main__":
    args: argparse.Namespace = parse_arguments()
    mdir: Optional[str] = args.mdir
    mbox: Optional[str] = args.mbox
    port: int = args.port

    mbox_path = Path(mdir) / mbox
    mbox_path.parent.mkdir(parents=True, exist_ok=True)

    asyncio.run(run_server("localhost", port, mbox_path))
