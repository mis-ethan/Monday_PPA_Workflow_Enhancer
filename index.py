from flask import Flask, request, jsonify
import requests
import os

app = Flask(__name__)


# Monday.com API
MONDAY_API_URL = "https://api.monday.com/v2"
API_KEY = os.getenv("MONDAY_API_KEY")
BOARD_ID = os.getenv("MONDAY_BOARD_ID")

