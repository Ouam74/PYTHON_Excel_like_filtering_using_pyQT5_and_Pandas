import os
import time
import pdb

def measure1(self):
    self.update_log.emit(2, "STARTING MEASUREMENT THREAD...OK")
    if self.isAborted == False:
            for i in range(100):
                if self.isAborted == False:
                    self.update_log.emit(2, "%d" %i)
                    time.sleep(0.1)
                    self.update_progressbar.emit(i)
                else:
                    break
    else:
        return (0)
    self.end_thread.emit()
    self.update_log.emit(2, "MEASUREMENT COMPLETED.\n")
    return (1)

def blink(self):
    while self.isrunning == True:
        self.update_blinkcolor.emit("background-color: green")
        time.sleep(0.5)
        self.update_blinkcolor.emit("background-color: red")
        time.sleep(0.5)