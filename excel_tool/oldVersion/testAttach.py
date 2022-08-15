# -*- coding:utf-8 -*-

# import VSCodeDebug
import ptvsd


def attach():
    host = "127.0.0.1"  # or "localhost"
    port = 12345
    print("Waiting for debugger attach at %s:%s ......" % (host, port))
    ptvsd.enable_attach(address=(host, port), redirect_output=True)
    ptvsd.wait_for_attach()
