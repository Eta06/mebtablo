import openpyxl
import json
import webbrowser
from flask import Flask, render_template, request, send_file, redirect

app = Flask(__name__)


def dosyakontrol():
    try:
        with open("tablo.xlsx", "r") as f:
            pass
    except FileNotFoundError:
        return False
    except:
        return False
    return True


def browserbaslat():
    webbrowser.open("http://127.0.0.1:6578")


def gunguncevirici(tablo, sheetname):
    data = okuyucu(tablo, sheetname)
    gunler = {
        "Pazartesi": {

        },
        "Salı": {

        },
        "Çarşamba": {

        },
        "Perşembe": {

        },
        "Cuma": {

        }
    }
    for i in gunler:
        # I = GÜNLER
        for j in data:
            # J = SINIFLAR
            if j not in gunler[i]:
                gunler[i][j] = {}
            for k in data[j][i]:
                # K = DERS NUMARALARI
                gunler[i][j][k] = data[j][i][k]

    with open("gunler.json", "w", encoding="utf-8") as f:
        json.dump(gunler, f, indent=4, ensure_ascii=False)
    return gunler


def guntoplayici(tablo):
    # Girilen Tablo Satırı İçerisinde Bulunan Bütün Dersleri Günlere Ayıran Fonksiyon
    # Output: gunler - map

    gunler = {
        "Pazartesi": {

        },
        "Salı": {

        },
        "Çarşamba": {

        },
        "Perşembe": {

        },
        "Cuma": {

        }
    }
    for e in range(3, 13):
        gunler["Pazartesi"][e - 2] = tablo[e]
    for e in range(13, 23):
        gunler["Salı"][e - 12] = tablo[e]
    for e in range(23, 33):
        gunler["Çarşamba"][e - 22] = tablo[e]
    for e in range(33, 43):
        gunler["Perşembe"][e - 32] = tablo[e]
    for e in range(43, 53):
        gunler["Cuma"][e - 42] = tablo[e]
    return gunler


def siniftoplayici(tablo):
    # Girilen Tablo İçerisinde Bulunan Bütün Sınıfların İsimlerini Toplayan Fonksiyon
    # Output: siniflar - map

    siniflar = {}
    for i in range(len(tablo)):
        siniflar[i] = tablo[i + 1][1]
    return siniflar


def okuyucu(datatable, sheetname):
    # Excel dosyasını aç
    wb = openpyxl.load_workbook(datatable)

    # "Dersliklerin Çarşaf Programı" adlı çalışma sayfasını seç
    sheet = wb[sheetname]

    mapdata = excel_to_map(sheet)

    new_map = delete_first_n_elements(mapdata, 3)

    map_sinif = siniftoplayici(new_map)

    siniflarin_ders_programi = {}
    for i in range(len(map_sinif)):
        siniflarin_ders_programi[map_sinif[i]] = guntoplayici(new_map[i + 1])

    return siniflarin_ders_programi


def excel_to_map(sheet):
    data = {}
    for row in sheet.iter_rows():
        data[row[0].value] = [cell.value for cell in row]
    return data


def delete_first_n_elements(my_map, n):
    return dict(list(my_map.items())[n:])


@app.route('/')
def index():
    if dosyakontrol():
        pass
    else:
        return redirect("/yenidosya")
    return render_template('index.html')


@app.route('/yenidosya', methods=["GET", "POST"])
def yenidosya():
    if dosyakontrol():
        pass
    else:
        return redirect("/yenidosya")
    if request.method == "POST":
        if "file" not in request.files:
            return redirect(request.url)
        file = request.files["file"]
        if file.filename == "":
            return redirect(request.url)
        if file.filename.split(".")[-1].lower() not in ["xlsx", "xls"]:
            return "Yalnızca .xlsx veya .xls dosyaları kabul edilmektedir!"
        file.save("tablo.xlsx")
        return render_template("index.html", message="Dosya başarılı bir şekilde yüklendi!")
    return render_template('yenidosya.html')


@app.route('/ogretmencarsaf')
def ogretmencarsaf():
    if dosyakontrol():
        pass
    else:
        return redirect("/yenidosya")
    return render_template("tablo.html", data=okuyucu("tablo.xlsx", "Öğretmenlerin Çarşaf Programı"), header="Tüm Öğretmenlerin Düzenlenmiş Çarşaf Ders Tablosu")


@app.route("/sinifcarsaf")
def sinifcarsaf():
    if dosyakontrol():
        pass
    else:
        return redirect("/yenidosya")
    return render_template("tablo.html", data=okuyucu("tablo.xlsx", "Sınıfların Çarşaf Programı"), header="Tüm Sınıfların Düzenlenmiş Çarşaf Ders Tablosu")


@app.route("/jsonkaydet", methods=["GET", "POST"])
def jsonkaydet():
    if dosyakontrol():
        pass
    else:
        return redirect("/yenidosya")
    sheetname = request.args.get("sheetname")
    sheetstyle = request.args.get("sheetstyle")
    if sheetname == "class":
        sheetname = "Sınıfların Çarşaf Programı"
    elif sheetname == "teacher":
        sheetname = "Öğretmenlerin Çarşaf Programı"
    else:
        return {"error": "Invalid sheetname!"}
    if sheetstyle == "gungun":
        veri = gunguncevirici("tablo.xlsx", sheetname)
    elif sheetstyle == "normal":
        veri = okuyucu("tablo.xlsx", sheetname)
    else:
        return {"error": "Invalid sheetstyle!"}
    with open(sheetname + '.json', 'w', encoding="utf-8") as f:
        json.dump(veri, f, indent=4, ensure_ascii=False)

    # JSON dosyasının kaydedildiği yol
    dosya_yolu = sheetname + ".json"

    # `send_file` fonksiyonunu kullanarak dosyayı gönderin
    return send_file(dosya_yolu, as_attachment=True)


@app.route("/ogretmengungun", methods=["GET", "POST"])
def ogretmengungun():
    if dosyakontrol():
        pass
    else:
        return redirect("/yenidosya")
    gun = request.args.get("gun")
    if gun is None:
        gun = "Pazartesi"

    return render_template("gungun.html", data=gunguncevirici("tablo.xlsx", "Öğretmenlerin Çarşaf Programı")[gun], gun=gun)


@app.route("/sinifgungun", methods=["GET", "POST"])
def sinifgungun():
    if dosyakontrol():
        pass
    else:
        return redirect("/yenidosya")
    gun = request.args.get("gun")
    if gun is None:
        gun = "Pazartesi"
    return render_template("sinifgungun.html", data=gunguncevirici("tablo.xlsx", "Sınıfların Çarşaf Programı")[gun], gun=gun)


if __name__ == '__main__':
    browserbaslat()
    app.run(port=6578)
