# chromeのセットアップ用
FROM public.ecr.aws/lambda/python:3.11 as build
# ドライバダウンロード
RUN yum install -y unzip && \
    curl -Lo "/tmp/chromedriver.zip" "https://chromedriver.storage.googleapis.com/103.0.5060.53/chromedriver_linux64.zip" && \
    curl -Lo "/tmp/chrome-linux.zip" "https://www.googleapis.com/download/storage/v1/b/chromium-browser-snapshots/o/Linux_x64%2F1002910%2Fchrome-linux.zip?alt=media" && \
    unzip /tmp/chromedriver.zip -d /opt/ && \
    unzip /tmp/chrome-linux.zip -d /opt/

FROM public.ecr.aws/lambda/python:3.11

RUN yum install atk cups-libs gtk3 libXcomposite alsa-lib \
    libXcursor libXdamage libXext libXi libXrandr libXScrnSaver \
    libXtst pango at-spi2-atk libXt xorg-x11-server-Xvfb \
    xorg-x11-xauth dbus-glib dbus-glib-devel -y

# 言語設定
ENV LANGUAGE=ja

# 必要なPythonライブラリインストール
COPY requirements.txt /var/task
RUN pip install -r /var/task/requirements.txt

# ダウンロードしてきたブラウザとドライバを配置
COPY --from=build /opt/chrome-linux /opt/chrome
COPY --from=build /opt/chromedriver /opt/

COPY src /var/task

CMD ["lambda_function.lambda_handler"]
