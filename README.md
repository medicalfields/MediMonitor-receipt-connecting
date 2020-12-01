# MediMonitorレセプトコンピューター連動ソフトウェア
MedicalFields株式会社のシステムであるMediMonitorと各社メーカーのレセプトコンピューター（以下レセコン）を連動させるためのソフトウェアのソースコードです。<br>
こちらをご確認していただけたら個人情報がサーバーへの送信の段階ですでに消されていることが分かります。<br>
また動作がシンプルなため他社レセコンソフトに不具合も起こりにくいと思われます。<br>
気になるようでしたら一度内部コードをご確認下さい。<br>
<br>
こちらのPythonプログラム(3.6.5)を32bitのwindows環境でpyinstallerを使用しEXE化し、inno setup complier でパッケージ化したものを当社のホームページにて配布しております。<br>
起動させるにはVisual Studio 2015 の Visual C++ 再頒布可能パッケージ が必要になります<br>
<br>
<h2>・Windows版について</h2>
<p>コンパイル済みパッケージのダウンロードし、インストールして下さい。<br>
ダウンロード→https://medicalfields.jp/mm_setup.exe<br>
※解説→https://medicalfields.jp/how_to_setup_mm/</p>
<br><br>
<h2>・Mac,Linux版について</h2>
<p>インストーラーはなく、各自で設定する必要があります。（専門的な知識が必要になります）
</p>

<br>
※Ubuntu18.04 LTS（ディスクトップ版）またはRaspbian OS 9.4(Raspberry Pi4 ModelB 4GB)での実行例<br><br>
Terminalでの実行<br>

<h6>1.ソフトウェアをダウンロード</h6>
curl -OL https://medicalfields.jp/medimonitor.tar.gz<br>
<h6>2.ダウンロードしたソフトを解凍する</h6>
tar zxvf medimonitor.tar.gz<br>
<h6>3.解凍したディレクトリに移動する</h6>
cd medimonitor<br>
<h6>4.解凍したディレクトリから連動ソフトを実行させる</h6>
python3 systray.py<br>
※ここでエラーが出る場合は<br>
sudo apt-get install -y python3-pip<br>
pip3 install pyqt5<br>
pip3 install objgraph<br>
sudo apt-get install python3-pyqt5<br>
をインストールして下さい<br>
<h6>5.連動ソフトが起動したら、共有先のSIPSフォルダをマウントする</h6>
sudo apt install -y cifs-utils<br>
sudo mkdir /mnt/sips<br>
sudo mount -t cifs //192.168.37.1/sips2 /mnt/sips -o user=sips,password=sips,iocharset=utf8<br>
※レセコンでの設定例（事前にレセコン側でSIPSの出力設定とフォルダ共有を行っている必要があります）<br>
IPアドレス：192.168.37.1<br>
共有フォルダ：sips2<br>
ユーザー：sips<br>
パスワード：sips<br>
<h6>6.マウントしたフォルダを設定ウィザードの連動先フォルダに入力しチェック、登録を行う</h6>
<h6>7.薬局IDとパスワードを入力し、ログインする</h6>
<br><br>

<h4>スタートアップ時にMediMonitor連動ソフトの自動起動の設定方法</h4>
※事前にpwdでMediMonitorレセコン連動ソフトウェアが存在するsystray.pyのフルパスを確認して下さい。<br>
例：python3 /home/user/medimonitor/systray.py<br>
・Ubuntuの場合<br>
「自動起動するアプリケーション」に、アプリ【python3 /home/user/medimonitor/systray.py】を登録<br>
・RaspberryPiの場合<br>
mkdir -p ~/.config/lxsession/LXDE-pi<br>
cp /etc/xdg/lxsession/LXDE-pi/autostart ~/.config/lxsession/LXDE-pi/<br>
vi ~/.config/lxsession/LXDE-pi/autostart<br>
で【pautostart】に<br>
python3 /home/user/medimonitor/systray.py<br>
を追加<br>
<br><br>

<h4>スタートアップ時にレセコンの共有フォルダへマウントさせる方法</h4>
sudo vi /etc/fstab<br>
で【fstab】に<br>
//192.168.37.1/sips2 /mnt/sips cifs username=sips,password=sips,iocharset=utf8,rw,defaults 0 0<br>
を追加<br>
<br><br>
<h5>※RaspberryPiでディスプレイ起動時にスタートアップが起動しない</h5>
sudo raspi-config<br>
3 Interface Options → Enable VNC<br>
2 Display Options →Resolution→1280x720<br>
<br><br>
<h5>※RaspberryPiでスタートアップ時にマウントされない</h5>
sudo vi /etc/rc.local<br>
で【rc.local】を<br>
#GS notes: a *minimum* of sleep 10 is required for the mount below to work on the Pi 3; it failed with sleep 5, but worked with sleep 10, sleep 15, and sleep 30<br>
sleep 20<br>
_IP=$(hostname -I) || true<br>
if [ "$_IP" ]; then<br>
printf "My IP address is %s\n" "$_IP"<br>
mount -a #GS: mount all drives in /etc/fstab<br>
fi<br>
exit 0<br>
に置き換え

<br><br>
設定方法について詳しくはhttps://medicalfields.jp/how_to_setup_mm/ を御覧下さい<br>
MediMonitor（メディモニター）についてはhttps://medicalfields.jp/medimonitor/<br>
MedicalFields株式会社についてはhttps://medicalfields.jp/ をご確認下さい<br>
