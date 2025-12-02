# app.py
import streamlit as st
from PIL import Image
import os
import numpy as np
import pickle

import tensorflow as tf
from tensorflow.keras.applications.resnet50 import ResNet50, preprocess_input
from tensorflow.keras.preprocessing import image as keras_image
from tensorflow.keras.models import Model


# ------------------------------------
# 기본 설정
# ------------------------------------
IMAGE_DB_DIR = "image_db"         # 상품 이미지 저장 폴더
METADATA_FILE = "metadata.pkl"    # 상품 정보 파일
TOP_K = 10                        # 유사도 기반 후보 개수


# ------------------------------------
# 이미지 임베딩 모델 (ResNet50)
# ------------------------------------
base_model = ResNet50(weights="imagenet", include_top=False, pooling="avg")
model = Model(inputs=base_model.input, outputs=base_model.output)


def get_embedding(img: Image.Image) -> np.ndarray:
    img = img.resize((224, 224))
    x = keras_image.img_to_array(img)
    x = np.expand_dims(x, axis=0)
    x = preprocess_input(x)

    feat = model.predict(x)
    feat = feat.flatten()
    feat = feat / np.linalg.norm(feat)  # normalize
    return feat


# ------------------------------------
# DB 로딩 함수
# ------------------------------------
@st.cache_data
def load_db():
    # 📌 폴더 없으면 자동 생성 (중요!)
    if not os.path.exists(IMAGE_DB_DIR):
        os.makedirs(IMAGE_DB_DIR)

    embeddings = {}
    metadata = {}

    # 메타데이터 로드 (없어도 에러 X)
    if os.path.exists(METADATA_FILE):
        with open(METADATA_FILE, "rb")_
