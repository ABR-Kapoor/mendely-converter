from flask import Flask, jsonify

app = Flask(__name__)


@app.get("/")
def home():
    return jsonify({
        "status": "ok",
        "app": "Mendeley Converter",
        "mode": "vercel-compatible api wrapper",
        "message": "This project is GitHub-ready and Vercel-compatible as a serverless API wrapper. The local Streamlit UI remains available via streamlit run app.py."
    })


@app.get("/health")
def health():
    return jsonify({"status": "healthy"})


if __name__ == "__main__":
    app.run(debug=True)
