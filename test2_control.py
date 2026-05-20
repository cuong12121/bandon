import argparse
import socket


def send_cut(host, port, barcode=None):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.connect((host, port))
    msg = "CUT" + (":" + barcode if barcode else "")
    s.sendall(msg.encode())
    s.close()


def send_stop(host, port):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.connect((host, port))
    s.sendall(b"STOP")
    s.close()


def main():
    parser = argparse.ArgumentParser(description="Send control commands to test2.py recorder")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=9999)
    parser.add_argument("--cut", help="Send CUT with optional barcode (e.g. --cut 1234)")
    parser.add_argument("--stop", action="store_true", help="Send STOP to recorder")
    args = parser.parse_args()

    if args.stop:
        send_stop(args.host, args.port)
        print("STOP sent")
    elif args.cut is not None:
        send_cut(args.host, args.port, args.cut)
        print(f"CUT sent (barcode={args.cut})")
    else:
        print("Nothing to do. Use --cut or --stop")


if __name__ == "__main__":
    main()
