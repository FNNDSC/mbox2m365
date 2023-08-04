import  asyncio
from    aiosmtpd import    controller
import  mailbox
import  os

class MboxMessageHandler:
    def __init__(self, mbox_path):
        self.mbox_path = mbox_path

    async def handle_DATA(self, server, session, envelope):
        # Save the message to an mbox file
        mbox = mailbox.mbox(self.mbox_path, create=True)
        mbox.add(envelope.content)
        mbox.flush()
        mbox.close()

        return '250 Message accepted for delivery'

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Mail sink server')
    parser.add_argument('--mdir',
                        required    = False,
                        help        = 'Output mbox directory',
                        default     = '/tmp/mail')
    parser.add_argument('--mbox',
                        required    = False,
                        help        = 'Output mbox file',
                        default     = 'messages.mbox')
    parser.add_argument('--port',
                        required    = False,
                        help        = 'Listener port',
                        default     = 22225)
    args = parser.parse_args()

    # Create the mbox directory if it doesn't exist
    print(f"Checking/creating mbox directory: {args.mdir}")
    os.makedirs(args.mdir, exist_ok=True)

    # Set up the asyncio event loop
    loop = asyncio.get_event_loop()

    # Start the SMTP server
    mbox_path = os.path.join(args.mdir, args.mbox)
    handler = MboxMessageHandler(mbox_path)
    controller = controller.Controller(handler, hostname='localhost', port=args.port)

    try:
        print(f"SMTP server is running. Listening on port {args.port}.")
        print(f"All email routed through this server will simply be saved to {args.mdir}/{args.mbox}.")
        controller.start()
#        loop.run_until_complete(controller.start())
        loop.run_forever()
    except KeyboardInterrupt:
        pass
    finally:
        controller.stop()
#        loop.run_until_complete(controller.stop())
        loop.close()

