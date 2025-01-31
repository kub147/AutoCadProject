import tkinter as tk
from tkinter import messagebox
import requests
import win32com.client
from shapely.wkb import loads
import pythoncom
import webbrowser
from pyproj import Transformer


# Funkcja do tworzenia tablicy współrzędnych w AutoCAD
def vArray(*args):
    """
    Funkcja tworząca tablicę VARIANT z argumentów.
    """
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, args)


# Funkcja do konwersji współrzędnych z układu PL-1992 (EPSG:2180) na WGS84 (EPSG:4326)
def konwertuj_wspolrzedne(x, y):
    """
    Konwertuje współrzędne z układu PL-1992 (EPSG:2180) na WGS84 (EPSG:4326).
    :param x: Współrzędna X w układzie PL-1992.
    :param y: Współrzędna Y w układzie PL-1992.
    :return: Krotka (latitude, longitude) w układzie WGS84.
    """
    transformer = Transformer.from_crs("EPSG:2180", "EPSG:4326", always_xy=True)
    longitude, latitude = transformer.transform(x, y)
    return latitude, longitude


# Funkcja do otwierania Google Maps z podanymi współrzędnymi
def otworz_google_maps(x, y):
    """
    Otwiera Google Maps z podanymi współrzędnymi.
    :param x: Współrzędna X w układzie PL-1992.
    :param y: Współrzędna Y w układzie PL-1992.
    """
    try:
        # Konwersja współrzędnych na format WGS84
        latitude, longitude = konwertuj_wspolrzedne(float(x), float(y))
        url = f"https://www.google.com/maps?q={latitude},{longitude}"
        webbrowser.open(url)
    except Exception as e:
        messagebox.showerror("Błąd", f"Nie udało się otworzyć Google Maps: {e}")


# Funkcja do pobierania danych działki na podstawie współrzędnych
def pobierz_dane_dzialki(wspolrzedne):
    url_api = f"https://uldk.gugik.gov.pl/?request=GetParcelByXY&xy={wspolrzedne}"
    try:
        odpowiedz = requests.get(url_api)
        odpowiedz.raise_for_status()

        odpowiedz_lines = odpowiedz.text.splitlines()

        if len(odpowiedz_lines) < 2:
            messagebox.showwarning("Brak danych", "Odpowiedź API jest niepełna.")
            return None

        status = odpowiedz_lines[0]
        if status != "0":
            messagebox.showwarning("Błąd", f"Błąd w odpowiedzi API: {status}")
            return None

        wspolrzedne_wkb = odpowiedz_lines[1]
        return {"status": status, "wspolrzedne_wkb": wspolrzedne_wkb}

    except requests.exceptions.RequestException as blad:
        messagebox.showerror("Błąd", f"Nie udało się pobrać danych działki: {blad}")
        return None


# Funkcja do pobierania danych gminy na podstawie współrzędnych
def pobierz_dane_commune(wspolrzedne):
    url_api = f"https://uldk.gugik.gov.pl/?request=GetCommuneByXY&xy={wspolrzedne}"
    try:
        odpowiedz = requests.get(url_api)
        odpowiedz.raise_for_status()

        odpowiedz_lines = odpowiedz.text.splitlines()

        if len(odpowiedz_lines) < 2:
            messagebox.showwarning("Brak danych", "Odpowiedź API jest niepełna.")
            return None

        status = odpowiedz_lines[0]
        if status != "0":
            messagebox.showwarning("Błąd", f"Błąd w odpowiedzi API: {status}")
            return None

        wspolrzedne_wkb = odpowiedz_lines[1]
        return {"status": status, "wspolrzedne_wkb": wspolrzedne_wkb}

    except requests.exceptions.RequestException as blad:
        messagebox.showerror("Błąd", f"Nie udało się pobrać danych commune: {blad}")
        return None


# Funkcja rysująca działkę w AutoCAD
def rysuj_dzialke_z_wkb(wkb):
    """
    Funkcja zamienia WKB na współrzędne i rysuje działkę w AutoCAD.
    :param wkb: Ciąg WKB reprezentujący geometrię działki.
    """
    try:
        # Dekodowanie WKB do obiektu Shapely
        geometria = loads(bytes.fromhex(wkb))

        # Sprawdzenie typu geometrii
        if not geometria.is_valid:
            raise ValueError("Geometria WKB jest nieprawidłowa.")
        if geometria.geom_type != "Polygon":
            raise ValueError(f"Obsługiwany jest tylko typ Polygon, a otrzymano: {geometria.geom_type}")

        # Pobranie współrzędnych z geometrii
        wspolrzedne = list(geometria.exterior.coords)
        print("Współrzędne działki:", wspolrzedne)  # Dodanie debugowania współrzędnych

        # Uruchomienie AutoCAD za pomocą pywin32
        acad = win32com.client.Dispatch("AutoCAD.Application")
        acad.Visible = True  # Sprawia, że AutoCAD będzie widoczny

        # Tworzenie nowego dokumentu
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace

        # Tworzenie listy współrzędnych w formacie (x1, y1, x2, y2, ...)
        punkty = []
        for punkt in wspolrzedne:
            punkty.append(punkt[0])  # dodajemy x
            punkty.append(punkt[1])  # dodajemy y

        # Przekazanie współrzędnych w tablicy do AutoCAD
        punkty_array = vArray(*punkty)

        # Dodawanie polilinii do modelu w AutoCAD
        linia = model_space.AddLightWeightPolyline(punkty_array)
        linia.Closed = True  # Zamknięcie polilinii, tworząc pełną obwódkę

        # Zmiana koloru linii (np. kolor 1 to czerwony)
        linia.Color = 1

        messagebox.showinfo("Sukces", "Działka została narysowana w AutoCAD na podstawie WKB!")
    except Exception as e:
        messagebox.showerror("Błąd", f"Nie udało się przetworzyć WKB lub narysować działki: {e}")
        print(f"Wyjątek: {e}")  # Dodanie logowania błędów


# Funkcja rysująca commune w AutoCAD
def rysuj_commune_z_wkb(wkb):
    """
    Funkcja zamienia WKB na współrzędne i rysuje commune w AutoCAD.
    :param wkb: Ciąg WKB reprezentujący geometrię commune.
    """
    try:
        # Dekodowanie WKB do obiektu Shapely
        geometria = loads(bytes.fromhex(wkb))

        # Sprawdzenie typu geometrii
        if not geometria.is_valid:
            raise ValueError("Geometria WKB jest nieprawidłowa.")

        if geometria.geom_type == "Polygon":
            # Dla zwykłego Polygona
            wspolrzedne = list(geometria.exterior.coords)
            print("Współrzędne commune (Polygon):", wspolrzedne)
            rysuj_poligon(wspolrzedne)

        elif geometria.geom_type == "MultiPolygon":
            # Dla MultiPolygon (kolekcja Polygonów)
            for poligon in geometria.geoms:
                wspolrzedne = list(poligon.exterior.coords)
                print("Współrzędne commune (MultiPolygon):", wspolrzedne)
                rysuj_poligon(wspolrzedne)

        else:
            raise ValueError(f"Nieobsługiwany typ geometrii: {geometria.geom_type}")

    except Exception as e:
        messagebox.showerror("Błąd", f"Nie udało się przetworzyć WKB lub narysować commune: {e}")
        print(f"Wyjątek: {e}")  # Dodanie logowania błędów


# Funkcja rysująca pojedynczy poligon w AutoCAD
def rysuj_poligon(wspolrzedne):
    """
    Funkcja do rysowania pojedynczego poligonu w AutoCAD.
    :param wspolrzedne: Lista współrzędnych w formie [(x1, y1), (x2, y2), ...].
    """
    try:
        # Uruchomienie AutoCAD za pomocą pywin32
        acad = win32com.client.Dispatch("AutoCAD.Application")
        acad.Visible = True  # Sprawia, że AutoCAD będzie widoczny

        # Tworzenie nowego dokumentu
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace

        # Tworzenie listy współrzędnych w formacie (x1, y1, x2, y2, ...)
        punkty = []
        for punkt in wspolrzedne:
            punkty.append(punkt[0])  # dodajemy x
            punkty.append(punkt[1])  # dodajemy y

        # Przekazanie współrzędnych w tablicy do AutoCAD
        punkty_array = vArray(*punkty)

        # Dodawanie polilinii do modelu w AutoCAD
        linia = model_space.AddLightWeightPolyline(punkty_array)
        linia.Closed = True  # Zamknięcie polilinii, tworząc pełną obwódkę

        # Zmiana koloru linii (np. kolor 3 to zielony)
        linia.Color = 3

        messagebox.showinfo("Sukces", "Commune zostało narysowane w AutoCAD na podstawie WKB!")
    except Exception as e:
        messagebox.showerror("Błąd", f"Nie udało się narysować poligonu: {e}")
        print(f"Wyjątek: {e}")  # Dodanie logowania błędów


# Funkcja obsługująca przesłanie formularza
def przeslij_formularz():
    lokalizacja = wpis_lokalizacja.get()

    if not lokalizacja:
        messagebox.showwarning("Błąd danych", "Wszystkie pola są wymagane!")
        return

    # Rozdzielamy współrzędne na x i y
    wspolrzedne = lokalizacja.strip()
    if wspolrzedne.count(',') == 1:
        x, y = wspolrzedne.split(',')
        x = x.strip()
        y = y.strip()

        # Pobranie danych działki i commune
        dane_dzialki = pobierz_dane_dzialki(f"{x},{y}")
        dane_commune = pobierz_dane_commune(f"{x},{y}")

        if dane_dzialki and dane_commune:
            # Wyświetlanie pobranych danych
            status_dzialki = dane_dzialki.get("status")
            wspolrzedne_wkb_dzialki = dane_dzialki.get("wspolrzedne_wkb")
            status_commune = dane_commune.get("status")
            wspolrzedne_wkb_commune = dane_commune.get("wspolrzedne_wkb")

            messagebox.showinfo("Sukces",
                                f"Pomyślnie pobrano dane!\nDziałka status: {status_dzialki}\nCommune status: {status_commune}")

            # Rysowanie działki w AutoCAD
            rysuj_dzialke_z_wkb(wspolrzedne_wkb_dzialki)

            # Rysowanie commune w AutoCAD
            rysuj_commune_z_wkb(wspolrzedne_wkb_commune)

            # Otwieranie Google Maps po zakończeniu rysowania
            otworz_google_maps(x, y)

        else:
            messagebox.showwarning("Brak danych", "Nie znaleziono danych dla podanych współrzędnych.")
    else:
        messagebox.showwarning("Błąd danych", "Wprowadź współrzędne w formacie: x,y")


# Tworzenie głównego okna Tkinter
okno = tk.Tk()
okno.title("Projektant Działki")
okno.geometry("600x300")
okno.configure(bg="#d6eaf8")

# Tworzenie etykiety dla lokalizacji
etykieta_lokalizacja = tk.Label(okno, text="Lokalizacja (współrzędne x,import tkinter as tk
from tkinter import messagebox
import requests
import win32com.client
from shapely.wkb import loads
import pythoncom
import webbrowser
from pyproj import Transformer


# Funkcja do tworzenia tablicy współrzędnych w AutoCAD
def vArray(*args):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, args)


def konwertuj_wspolrzedne(x, y):
    transformer = Transformer.from_crs("EPSG:2180", "EPSG:4326", always_xy=True)
    longitude, latitude = transformer.transform(x, y)
    return latitude, longitude


def otworz_google_maps(x, y):
    try:
        # Konwersja współrzędnych na format WGS84
        latitude, longitude = konwertuj_wspolrzedne(float(x), float(y))
        url = f"https://www.google.com/maps?q={latitude},{longitude}"
        webbrowser.open(url)
    except Exception as e:
        messagebox.showerror("Błąd", f"Nie udało się otworzyć Google Maps: {e}")


def pobierz_dane_dzialki(wspolrzedne):
    url_api = f"https://uldk.gugik.gov.pl/?request=GetParcelByXY&xy={wspolrzedne}"
    try:
        odpowiedz = requests.get(url_api)
        odpowiedz.raise_for_status()

        odpowiedz_lines = odpowiedz.text.splitlines()

        if len(odpowiedz_lines) < 2:
            messagebox.showwarning("Brak danych", "Odpowiedź API jest niepełna.")
            return None

        status = odpowiedz_lines[0]
        if status != "0":
            messagebox.showwarning("Błąd", f"Błąd w odpowiedzi API: {status}")
            return None

        wspolrzedne_wkb = odpowiedz_lines[1]
        return {"status": status, "wspolrzedne_wkb": wspolrzedne_wkb}

    except requests.exceptions.RequestException as blad:
        messagebox.showerror("Błąd", f"Nie udało się pobrać danych działki: {blad}")
        return None


def pobierz_dane_commune(wspolrzedne):
    url_api = f"https://uldk.gugik.gov.pl/?request=GetCommuneByXY&xy={wspolrzedne}"
    try:
        odpowiedz = requests.get(url_api)
        odpowiedz.raise_for_status()

        odpowiedz_lines = odpowiedz.text.splitlines()

        if len(odpowiedz_lines) < 2:
            messagebox.showwarning("Brak danych", "Odpowiedź API jest niepełna.")
            return None

        status = odpowiedz_lines[0]
        if status != "0":
            messagebox.showwarning("Błąd", f"Błąd w odpowiedzi API: {status}")
            return None

        wspolrzedne_wkb = odpowiedz_lines[1]
        return {"status": status, "wspolrzedne_wkb": wspolrzedne_wkb}

    except requests.exceptions.RequestException as blad:
        messagebox.showerror("Błąd", f"Nie udało się pobrać danych commune: {blad}")
        return None


def rysuj_dzialke_z_wkb(wkb):
    """
    Funkcja zamienia WKB na współrzędne i rysuje działkę w AutoCAD.
    :param wkb: Ciąg WKB reprezentujący geometrię działki.
    """
    try:
        # Dekodowanie WKB do obiektu Shapely
        geometria = loads(bytes.fromhex(wkb))

        # Sprawdzenie typu geometrii
        if not geometria.is_valid:
            raise ValueError("Geometria WKB jest nieprawidłowa.")
        if geometria.geom_type != "Polygon":
            raise ValueError(f"Obsługiwany jest tylko typ Polygon, a otrzymano: {geometria.geom_type}")

        # Pobranie współrzędnych z geometrii
        wspolrzedne = list(geometria.exterior.coords)
        print("Współrzędne działki:", wspolrzedne)  # Dodanie debugowania współrzędnych

        acad = win32com.client.Dispatch("AutoCAD.Application")
        acad.Visible = True  # Sprawia, że AutoCAD będzie widoczny

        doc = acad.ActiveDocument
        model_space = doc.ModelSpace

        punkty = []
        for punkt in wspolrzedne:
            punkty.append(punkt[0])
            punkty.append(punkt[1])

        punkty_array = vArray(*punkty)

        linia = model_space.AddLightWeightPolyline(punkty_array)
        linia.Closed = True

        linia.Color = 1

        messagebox.showinfo("Sukces", "Działka została narysowana w AutoCAD na podstawie WKB!")
    except Exception as e:
        messagebox.showerror("Błąd", f"Nie udało się przetworzyć WKB lub narysować działki: {e}")
        print(f"Wyjątek: {e}")  # Dodanie logowania błędów


def rysuj_commune_z_wkb(wkb):
    """
    Funkcja zamienia WKB na współrzędne i rysuje commune w AutoCAD.
    :param wkb: Ciąg WKB reprezentujący geometrię commune.
    """
    try:
        geometria = loads(bytes.fromhex(wkb))

        if not geometria.is_valid:
            raise ValueError("Geometria WKB jest nieprawidłowa.")

        if geometria.geom_type == "Polygon":
            wspolrzedne = list(geometria.exterior.coords)
            print("Współrzędne commune (Polygon):", wspolrzedne)
            rysuj_poligon(wspolrzedne)

        elif geometria.geom_type == "MultiPolygon":
            for poligon in geometria.geoms:
                wspolrzedne = list(poligon.exterior.coords)
                print("Współrzędne commune (MultiPolygon):", wspolrzedne)
                rysuj_poligon(wspolrzedne)

        else:
            raise ValueError(f"Nieobsługiwany typ geometrii: {geometria.geom_type}")

    except Exception as e:
        messagebox.showerror("Błąd", f"Nie udało się przetworzyć WKB lub narysować commune: {e}")
        print(f"Wyjątek: {e}")  # Dodanie logowania błędów


def rysuj_poligon(wspolrzedne):
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        acad.Visible = True  # Sprawia, że AutoCAD będzie widoczny

        doc = acad.ActiveDocument
        model_space = doc.ModelSpace

        punkty = []
        for punkt in wspolrzedne:
            punkty.append(punkt[0])
            punkty.append(punkt[1])

        punkty_array = vArray(*punkty)

        linia = model_space.AddLightWeightPolyline(punkty_array)
        linia.Closed = True
        linia.Color = 3

        messagebox.showinfo("Sukces", "Commune zostało narysowane w AutoCAD na podstawie WKB!")
    except Exception as e:
        messagebox.showerror("Błąd", f"Nie udało się narysować poligonu: {e}")
        print(f"Wyjątek: {e}")  # Dodanie logowania błędów


def przeslij_formularz():
    lokalizacja = wpis_lokalizacja.get()

    if not lokalizacja:
        messagebox.showwarning("Błąd danych", "Wszystkie pola są wymagane!")
        return

    wspolrzedne = lokalizacja.strip()
    if wspolrzedne.count(',') == 1:
        x, y = wspolrzedne.split(',')
        x = x.strip()
        y = y.strip()

        dane_dzialki = pobierz_dane_dzialki(f"{x},{y}")
        dane_commune = pobierz_dane_commune(f"{x},{y}")

        if dane_dzialki and dane_commune:
            status_dzialki = dane_dzialki.get("status")
            wspolrzedne_wkb_dzialki = dane_dzialki.get("wspolrzedne_wkb")
            status_commune = dane_commune.get("status")
            wspolrzedne_wkb_commune = dane_commune.get("wspolrzedne_wkb")

            messagebox.showinfo("Sukces",
                                f"Pomyślnie pobrano dane!\nDziałka status: {status_dzialki}\nCommune status: {status_commune}")

            rysuj_dzialke_z_wkb(wspolrzedne_wkb_dzialki)

            rysuj_commune_z_wkb(wspolrzedne_wkb_commune)

            otworz_google_maps(x, y)

        else:
            messagebox.showwarning("Brak danych", "Nie znaleziono danych dla podanych współrzędnych.")
    else:
        messagebox.showwarning("Błąd danych", "Wprowadź współrzędne w formacie: x,y")


okno = tk.Tk()
okno.title("Rysowanie Działki")
okno.geometry("600x300")
okno.configure(bg="#d6eaf8")

etykieta_lokalizacja = tk.Label(okno, text="Lokalizacja (współrzędne x,y):", bg="#d6eaf8", fg="#2c3e50",
                                font=("Arial", 12))
etykieta_lokalizacja.grid(row=0, column=0, padx=10, pady=10, sticky="e")

wpis_lokalizacja = tk.Entry(okno, width=40, font=("Arial", 12))
wpis_lokalizacja.grid(row=0, column=1, padx=10, pady=10)

wpis_lokalizacja.insert(0, '565186.44,244004.32')  # domylnie błonia

przycisk_przeslij = tk.Button(okno, text="Prześlij", command=przeslij_formularz, bg="#2980b9", fg="white",
                              font=("Arial", 14, "bold"))
przycisk_przeslij.grid(row=1, column=0, columnspan=2, pady=20)

okno.mainloop()y):", bg="#d6eaf8", fg="#2c3e50",
                                font=("Arial", 12))
etykieta_lokalizacja.grid(row=0, column=0, padx=10, pady=10, sticky="e")

# Tworzenie pola tekstowego dla lokalizacji
wpis_lokalizacja = tk.Entry(okno, width=40, font=("Arial", 12))
wpis_lokalizacja.grid(row=0, column=1, padx=10, pady=10)

# Ustawienie domyślnej wartości dla pola tekstowego
wpis_lokalizacja.insert(0, '565186.44,244004.32')  # To jest poprawna metoda dla Entry

# Przycisk przesyłania
przycisk_przeslij = tk.Button(okno, text="Prześlij", command=przeslij_formularz, bg="#2980b9", fg="white",
                              font=("Arial", 14, "bold"))
przycisk_przeslij.grid(row=1, column=0, columnspan=2, pady=20)

# Uruchomienie pętli głównej Tkinter
okno.mainloop()
