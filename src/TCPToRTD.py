#  TCPToRTD.py - Pyuno/LO bridge to implement new functions for LibreOffice Calc
#
#  license: GNU LGPL
#
#  This library is free software; you can redistribute it and/or
#  modify it under the terms of the GNU Lesser General Public
#  License as published by the Free Software Foundation; either
#  version 3 of the License, or (at your option) any later version.

import inspect
import logging
import os
import pathlib
import platform
import ssl
import sys
import time
import threading
import socket
import re
import datetime
from functools import wraps
from importlib import util

import unohelper
from org.tcptortd.getinfo import TCPToRTD

basedir = os.path.join(str(pathlib.Path.home()), '.RTD-extension')
os.makedirs(basedir, exist_ok=True)

logging.basicConfig(
    handlers=[logging.FileHandler(filename=os.path.join(basedir, 'extension.log'), encoding='utf-8', mode='a+')],
    format="%(asctime)s %(name)s %(levelname)s %(message)s",
    level=logging.INFO)

# Add current directory to import path
current_dir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

implementation_name = "org.tcptortd.getinfo.python.TCPToRTDImpl"  # as defined in TCPToRTD.xcu
implementation_services = ("com.sun.star.sheet.AddIn",)

def profile(fn):
    @wraps(fn)
    def with_profiling(*args, **kwargs):
        start = time.perf_counter()
        r = fn(*args, **kwargs)
        elapsed = time.perf_counter() - start

        with open(os.path.join(basedir, 'trace.log'), "a+", encoding="utf-8") as text_file:
            print(
                f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')} {fn.__name__} *args={args[1:]} r='{r}' {(1000 * elapsed):.3f} ms",
                file=text_file)

        return r

    return with_profiling

# This is a TCP listener (dedicated thread) that accepts a connection and parses incoming data
# with pairs of "<key|value>". Each pair is sent to the registered key_value_handler_func
class TCPToRTDListener(threading.Thread):
    def __init__(self, key_value_handler_func, port, host = "0.0.0.0"):
        threading.Thread.__init__(self)
        self.port = port
        self.host = host
        self.key_value_handler_func = key_value_handler_func

    def run(self):
        while True:
            server_socket = socket.socket()
            server_socket.bind((self.host, self.port))
            server_socket.listen(1)
            logging.info("Listening on " + str(self.host) + ":" + str(self.port))
            conn, address = server_socket.accept()  # Wait for a new connection
            logging.info("Connection from: {}".format(address))
            self.handle_connection(conn)

    def handle_connection(self, conn):
        pattern = re.compile(r"<(.+?)\|(.+?)>") # <key|value>
        str_data = ""

        while True:
            # receive data stream. it won't accept data packet greater than 1024 bytes
            data = conn.recv(1024).decode()
            if not data:
                # if data is not received break
                break
            
            str_data = str_data + str(data) # Concat leftover data from prev. block of data

            # find all matches to groups
            for match in pattern.finditer(str_data):
                self.key_value_handler_func(match.group(1), match.group(2))

            str_data = re.sub(r"<.+>", "", str_data)

        conn.close()  # close the connection

# This class represents a single RTD result which will change as data arrives
# Implementation inspired by https://wiki.documentfoundation.org/Documentation/DevGuide/Spreadsheet_Documents#Spreadsheet_Add-Ins
from com.sun.star.sheet import XVolatileResult
from com.sun.star.sheet import XResultListener
from com.sun.star.sheet import ResultEvent

class RTDResult(unohelper.Base, XVolatileResult):
    @profile
    def __init__(self, name, value = None):
        logging.info("Initializing RTDResult for key <{}>".format(name))
        self.listeners = set() # LO Calc listeners
        self.name = name
        self.value = value

    # adds a listener to be notified when a new value is available.
    def addResultListener(self, aListener):
        self.listeners.add(aListener)
        aListener.modified(self.getResult())

    # removes the specified listener.
    def removeResultListener(self, aListener):
        self.listeners.remove(aListener)

    def getResult(self):
        #logging.info("Starting getResult()")
        aEvent = ResultEvent()
        aEvent.Value = str(self.value)
        aEvent.Source = self
        return aEvent

    def modify(self, aNewValue):
        #logging.info("Starting modify()")
        self.value = aNewValue
        aEvent = self.getResult()
        for listener in self.listeners:
            #logging.info("Notifying listener " + str(listener))
            listener.modified(aEvent)

# This class implements the user-defined function (UDF) RTD() of this extension
class TCPToRTDImpl(unohelper.Base, TCPToRTD):
    """Define the main class for the TCPToRTD extension """

    @profile
    def __init__(self, ctx):
        logging.info("Initializing TCPToRTDImpl...")
        self.ctx = ctx
        self.dict = {} # string key -> RTDResult
        self.dict["__count.keys"] = RTDResult("__count.keys", 0)
        self.dict["__count.updates"] = RTDResult("__count.updates", 0)
        self.port = 13000

    def RTD(self, key):

        #logging.info("RTD() called from Calc with key <{}>".format(key))

        # Make sure key is legal
        try:
            if type(key) == tuple or not key:
                return None

            key = str(key).strip()

        except Exception as ex:
            logging.exception("Exception")
            return None

        # Return RTDResult for requested key
        if (not key in self.dict):
            self.dict[key] = RTDResult(key)
            countKeysRTDResult = self.dict["__count.keys"]
            countKeysRTDResult.modify(len(self.dict))
            
            # Start TCP socket listener if not running yet
            if (key == "__start"):
                self.start_listener(self.port)
                self.dict[key].modify(self.port)
        
        return self.dict[key]

    @profile
    def start_listener(self, port):
        logging.info("Starting TCP listener on port {}...".format(port))
        self.listener = TCPToRTDListener(self.handle_key_value_immediately, port)
        self.listener.start()
        return "TCP Listener started on port {}".format(port)

    def handle_key_value_immediately(self, key, value):
        # logging.info("Handling key and value: <{}> --> <{}>".format(key, value))
        countUpdatesRTDResult = self.dict["__count.updates"]
        countUpdatesRTDResult.modify(countUpdatesRTDResult.value + 1)
        if (countUpdatesRTDResult.value % 1000 == 0):
            logging.info("Handled {} items".format(countUpdatesRTDResult.value))

        aRTDResult = self.dict.get(key, None)
        if (not aRTDResult is None):
            aRTDResult.modify(value)

def createInstance(ctx):
    return TCPToRTDImpl(ctx)

# python loader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(createInstance, implementation_name, implementation_services, )

# For testing purposes
if __name__ == '__main__':

    aTCPToRTD = TCPToRTDImpl("nocontext")
    aTCPToRTD.start_listener(13000)

    while True:
      print("Listening...({})".format(aTCPToRTD.counter_updates))
      time.sleep(10)