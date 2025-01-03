import sys
import os

# Proje kök dizinini Python path'ine ekle
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from views.ana_pencere import AnaPencere

if __name__ == "__main__":
    app = AnaPencere()
    app.mainloop() 