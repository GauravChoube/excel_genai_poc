from flask import Flask, jsonify
import pandas as pd

app = Flask(__name__)

@app.route('/data')
def get_data():
    data = {'A': [1, None, 3], 'B': [None, 5, 6]}
    df = pd.DataFrame(data)
    result_dict = df.to_dict(orient='records')
    return jsonify(result_dict)

if __name__ == '__main__':
    app.run(debug=True)