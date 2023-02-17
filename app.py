import pandas as pd
from flask import Flask, render_template, request

app = Flask(__name__)

# carrega a planilha Excel
df = pd.read_excel('Planilha.xlsx', engine='openpyxl')


# define a rota e a função para localizar a célula
@app.route('/')
def get_data():
    df = pd.read_excel('Planilha.xlsx', engine='openpyxl')

    return df.to_json()

@app.route('/busca', methods=['GET', 'POST'])
def busca():
    if request.method == 'POST':
        search_term = request.form['search_term']

        results = df.loc[df['Nome'].str.contains(search_term), ['Nome', 'Mesa']]
        if not results.empty:
            message = ""
            for nome, mesa in results.to_dict('split')['data']:
                message += f'O convidado "{nome}" está na mesa <button class="table-button" value="{mesa}">{mesa}</button><br>'
        else:
            message = f'O convidado "{search_term}" não foi encontrado'

        return render_template("Busca.html", message=message)
    else:
        return render_template("Busca.html")

@app.route('/mesa/<mesa>')
def mesa(mesa):
    results = df.loc[df['Mesa'] == int(mesa), 'Nome']
    message = f'Os convidados na mesa {mesa} são:<br>'
    for nome in results:
        message += f'{nome}<br>'
    return render_template("Mesa.html", mesa_value=message)



# roda o aplicativo
# http://127.0.0.1:5000/busca

if __name__ == '__main__':
    app.run(debug=True)
