from flask import Flask, render_template,request,g,Request
from test_table import class_tabel

tabel_name = class_tabel.app
print("\n",tabel_name().result_umi)
sora = tabel_name().result_sora
umi = tabel_name().result_umi
hare = tabel_name().result_hare
sonota = tabel_name().result_sonota
a = [1,2,3,4,5,"Hello"]
b = [6,7,8,9,"Hello"]
c = [11,12,13,14,15,"Hello"]

app = Flask(__name__)

@app.route('/')
def home():

    return render_template('tabel.html',sora=sora,umi=umi,hare=hare,sonota=sonota)


def main():
    app.debug=True
    app.run()


if __name__ == "__main__":
    main()
