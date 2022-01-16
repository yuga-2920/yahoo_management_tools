import eel
import sys
import desktop
import yahuoku_shipping

app_name="web"
end_point="index.html"
size=(380, 300)

@ eel.expose
def main():
    yahuoku_shipping.main()
    
desktop.start(app_name,end_point,size)