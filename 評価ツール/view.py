import eel
import sys
import desktop
import yahuoku_evaluation

app_name="web"
end_point="index.html"
size=(280,650)

@ eel.expose
def main():
    yahuoku_evaluation.main()
    # sys.exit(1)
    
desktop.start(app_name,end_point,size)
# sys.exit(1)
