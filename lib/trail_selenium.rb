#! ruby -EWindows-31J
# -*- mode:ruby; coding:Windows-31J -*-

#= Trail4You Selenium拡張モジュール
#
#
require "trail_selenium/version"


require 'rubygems'
require 'selenium-webdriver'
require './excel'
require 'win32/clipboard'
require 'fileutils'
require 'date'
require 'fastimage'
require 'uri'
require 'open-uri'
require 'openssl'

# SSL証明書でエラーが起こるのを防ぐ為に無視する
# 証明書に問題が無ければコメントアウトしてください
#
OpenSSL::SSL::VERIFY_PEER = OpenSSL::SSL::VERIFY_NONE  


# STDOUT.sync = true;


#-------------------------------------------------------------------

#= Trail_Selenium
#
#Authors:: Mt.Trail
#Version:: 1.0 2016/7/24 Mt.Trail
#Copyright:: Copyrigth (C) Mt.Trail 2016 All rights reserved.
#License:: GPL version 2
#
#==目的
# Seleniumでデータ収集するためのクラス
#*   Excelクラスも利用する。
#*   Windows用です。
#*   パスは'/'区切りで扱います。'\\'では有りません。
#
class Trail_Selenium
  attr_accessor :driver,:wait
  attr_accessor :report_book,:report_sheet,:report_line_no
  
  #=== 初期化 wait時間(秒)を指定する。
  def initialize (wait_time = 10)
    @driver = Selenium::WebDriver.for :firefox
    @wait = Selenium::WebDriver::Wait.new(:timeout => wait_time) # seconds
    @report_book = nil
    @report_sheet = nil
    @report_line_no = 1
  end

  #=== ログイン
  #    引数で設定値の配列の配列を渡す、一番最後はsubmitボタンの情報(設定値なし)
  #    各配列要素は下記の形式
  #    [属性名シンボル,属性の値,設定値]　又は　[:xpath, 'xpath指定',設定値]
  #    例 : login([[:name,'UserName','LoginName'],[:name,'Password','password'],[:name,'Submit']])
  #
  def login (param)
    param.each_with_index do |p,i|
      if i < (param.size - 1)
        @driver.find_element(p[0], p[1]).send_keys(p[2])
      else
        @driver.find_element(p[0], p[1]).click
      end
    end
  end

  
  #-------------------------------------------------------------------
  
  #=== データ書き込み用のExcelを指定
  #  target : openするExcelファイルのパス
  #  tenplate : テンプレートのexcelファイルのパス、これをtargetにコピーしてからopenする。
  #
  #   テンプレートを指定するとそれをコピーして使用する。
  #   ブロックで処理内容を受け取る
  #
  def open_excel (target,tenplate = '')
    @target_excel = target

    if tenplate != ''
      FileUtils.cp(  tenplate, @target_excel)
    end

    openExcelWorkbook(@target_excel) do |book|
      @report_book = book
      yield book
    end
  end

  #-------------------------------------------------------------------

  #=== Excelシートのオープン
  #  book : openされたexcelオブジェクト
  #  sheet_name : シート番号またはシート名
  #
  #   ブックとシート名を指定する。
  #   シート名が数値の場合シートの番号と見なされる
  #   Excelのシートオブジェクトを返す
  #
  #   例
  #     ts = Trail_Selenium.new
  #     ts.open_excel('report.xls','report_tenplate.xls') do |book|
  #       sheet = open_report_sheet(book,'Report_Sheet')
  #         :
  #       book.save
  #     end
  #
  def open_sheet( book, sheet_name )
    sheet = book.Worksheets.Item(sheet_name)
    sheet.extend Worksheet
    sheet
  end

  #=== レポート用のExcelシートのオープン
  #  book : openされたexcelオブジェクト
  #  sheet_name : シート番号またはシート名
  #
  # @report_sheetを設定する
  # Excelに書き込む場合、こちらを指定すると書き込み関数呼び出し時にパラメータを減らせる。
  #
  def open_report_sheet( book, sheet_name )
    @report_sheet = open_sheet( book, sheet_name )
    @report_sheet
  end


  #-------------------------------------------------------------------

  #=== コンソールへの表示とReportシートへの記録
  #   offset: はコンソール出力時の左マージンとして使用される。
  #         : またExcelシートの場合、何カラム目からデータをセットするかの指定となる。
  #   t     : 出力する文字列の配列を指定する。コンソールとExcelシートに出力される。
  #   sheet : デフォルトの@report_sheet以外のシートに出力するときハッシュで指定する。:sheet => other_sheet
  #   line_no : 出力の行番号がデフォルトの@report_line_no以外のときハッシュで指定する。 :line_no => 2
  #
  #   出力するExcelシートと出力行はハッシュで指定する。指定されない場合、最後にopen_report_sheetで開いたシートが使われる。
  #   出力する行は指定されない場合 @report_line_noが使用される。
  #
  #   出力後はsheetが指定されていない場合 @report_line_noは + 1 される。
  #   出力文字列にカンマ等を含まないという制限条件はあるがoffset=0のコンソール出力をファイルにリダイレクトするとCSVファイルとなる。
  #
  #   注意 : Excelへ出力する場合open_excelのブロック内で利用されなければならない。
  #
  def disp_msg_array(offset,t=[''],sheet: nil,line_no: nil )
    print '  '*offset + t.map{|x| x.to_s}.join(', ') + "\n"
    
    sheet = @report_sheet if !sheet
    line = @report_line_no if !line_no
    
    if sheet
      t.each_with_index do |tt,i|
        sheet[line, offset+i+1] = tt
      end
      @report_line_no += 1 if !line_no
    end
  end
  
  #-------------------------------------------------------------------

  #=== xpathで指定された画像エレメントから画像をコピー機能を使用し、クリップボード経由でファイルに落とす。
  # 動的に生成される画像を保存するときに使用する。
  # 右クリックで画像をコピーメニューが出ないものには使用できない。
  #
  #   node     : xpathの開始ノード
  #   xp       : 画像を指定するxpath
  #   filename : 書き込む画像ファイル名
  #   wait_mode: エレメントの出現を待つとき true を指定 :wait_mode=>true
  #
  def get_picture_via_clipboard(filename,xp,node: nil,wait_mode:nil)
    if wait_mode
      img = find_element_until(xp,:node => node)
    else
      img = find_element(xp,:node => node)
    end
    
    if img
      @driver.action.context_click(img).send_keys('Y').perform
      if Win32::Clipboard.format_available?(Win32::Clipboard::DIB)
        File.open(filename,'wb') do |f|
          f.write Win32::Clipboard.data(Win32::Clipboard::DIB)
        end
      end
    end
    img
  end

  #-------------------------------------------------------------------

  #=== xpathで指定された画像エレメント(imgタグ)のURLから画像をファイルに落とす。
  #   node     : xpathの開始ノード
  #   xp       : 画像を指定するxpath
  #   pathname : 画像ファイルを書き込むフォルダパス
  #   wait_mode: エレメントの出現を待つとき true を指定 :wait_mode=>true
  #   rename   : ファイル名を元の名前から書き換えるとき指定 :rename => 'newname.jpg'
  #            : 指定されなければsrc属性に指定されたファイル名が使用される。
  #
  #   <return> : 画像ファイルパス or nil
  #
  def get_picture(pathname,xp,node: nil,wait_mode:nil,rename:nil)
    if wait_mode
      img = find_element_until(xp,:node => node)
    else
      img = find_element(xp,:node => node)
    end
    
    savefile = nil

    if img
      pathname += '/' if (pathname != '') and (pathname[-1] != '/')
      url = img[:src]
      if rename
        savefile = pathname + rename
      else
        filename = File.basename(url)
        savefile = pathname + filename
      end

      open(savefile,'wb') do |wf|
        open(url) do |rf|
          wf.write( rf.read )
        end
      end
    end
    savefile
  end

  #-------------------------------------------------------------------

  #=== Excelに画像貼り付け
  #  filename : 貼り付ける画像ファイル(Excel内に取り込まれる)
  #  cx,cy : 貼り付け位置のカラム(cx)と行(cy) 1始まりの値
  #  sh,sw : 画像の貼り付けドットサイズ 高さ(sh) 幅(sw)
  #  sheet : 貼り付けるシートオブジェクト、指定無しの場合@report_sheet
  #  fit_x : カラム幅をswに合わせる。
  #  fit_y : 行の高さをshに合わせる。
  #
  def add_picture_to_excel(filename,cy,cx,sw,sh,sheet: nil,fit_x: nil,fit_y: nil)
    sheet = @report_sheet if ! sheet
    if sheet
      r = sheet.Range(sheet.r_str(cy,cx))
      sheet.Shapes.AddPicture(filename.gsub("\/","\\"),false,true, r.Left.to_i, r.Top.to_i, 0.75*sw, 0.75*sh)
      sheet.set_width(cy,cx,0.118*sw)  if fit_x
      sheet.set_height(cy,cx,0.75*sh)  if fit_y
    end
  end


  #=== セレクトBOX選択
  #   xp : セレクトエレメントを指定するxpath
  #   tx : 選択する文字列の内容
  #
  def select_by_text(xp,tx)
    select = Selenium::WebDriver::Support::Select.new( @wait.until{@driver.find_element(:xpath,xp)} )
    select.select_by(:text,tx.encode('UTF-8'))
  end
  
  
  #=== エレメント探索
  #  xp   : セレクトエレメントを指定するxpath
  #  node : 途中の要素からの場合、その要素オブジェクトを指定する :node => element
  #
  #   見つからないときにはnilを返す。
  #
  def find_element(xp,node: nil)
    begin
      if node
        link = node.find_element(:xpath,xp)
      else
        link = @driver.find_element(:xpath,xp)
      end
    rescue
      link = nil
    end
    link
  end

  #=== エレメント探索(見つかるまで待つ)
  #  xp   : セレクトエレメントを指定するxpath
  #  node : 途中の要素からの場合、その要素オブジェクトを指定する :node => element
  #
  #   見つからないときにはnilを返す。
  #
  def find_element_until(xp,node: nil)
    begin
      if node
        link = @wait.until{node.find_element(:xpath,xp)}
      else
        link = @wait.until{@driver.find_element(:xpath,xp)}
      end
    rescue
      link = nil
    end
    link
  end
  
  #=== Selenium 終了
  #
  def close
    @driver.quit()
  end

end   ## --- end of Trail_Selenium class ---


if $0 == __FILE__

require 'date'

  #
  # Twitterのメッセージと画像を取る例です。
  #
  ts = Trail_Selenium.new    # クラスの初期化
  driver = ts.driver
  driver.navigate.to( 'https://twitter.com/3190')  # 穂高山荘さんのページ

  # ログインのダイアログが出ていたら消す
  if ts.find_element('//input[@name="session[username_or_email]"]').displayed?
    ts.find_element('//small[text() = "アカウントをお持ちの場合"]').click
  end

  # 最初の画像を pic1.jpg という名前で保存してみる。
  ts.get_picture('','//div[@class="AdaptiveMedia-photoContainer js-adaptive-photo "]/img',:rename=>'Hotaka.jpg')
  
  # 現在読み込まれているpostを取り出す
  posts = driver.find_elements(:xpath,'//div[@class="stream"]/ol/li')
  print "Post数 = #{posts.size}\n"
  
  # 保存用のExcelファイルをテンプレートからコピーして開く
  ts.open_excel("Hotakasansou_#{DateTime.now.strftime("%Y%m%d%H%M%S")}.xls",'Report_tenplate.xls') do |book|

    # 書き込み用のシートを開く
    sheet = ts.open_report_sheet(book,'Report')

    ts.report_line_no = 2      # 書込み開始は2行目から
    sheet.set_width(1,1,11.8)  # セル(1,1)を100pixcelのサイズにする。
    sheet.set_height(1,1,75)   # Y: 100pixcel = 75 point
    sheet.set_width(1,3,11.8)  # X: 100pixcel = 11.8 文字  画像用のセルも幅を設定

    # 読み込んだpostを順に処理
    posts.each do |post|
      
      # ツイートした人の名前
      name = post.find_element(:xpath,'.//div[@class="stream-item-header"]/a/strong').text()
      
      # メッセージ内容
      msg  = post.find_element(:xpath,'.//p').text()
      
      # コンソールとExcelへ出力
      ts.disp_msg_array(0,[name,msg])

      # 絶対パスでのファイル名を作成
      temp_file = File.expand_path('../temp.jpg',__FILE__)

      # 画像があるなら取り出し
      if ts.get_picture_via_clipboard(temp_file, './/div[@class="content"]/div[3]//img',:node=>post)
        # ピクセルサイズを調べる
        sz = FastImage.size(temp_file)
        
        # Excelに貼り付けreport_line_noは既に次の行に行っているので補正する
        # 画像サイズは幅が100pixcelになるように補正
        # fit_yをtrueにしてセルの高さを画像の高さに合わせる。
        ts.add_picture_to_excel(temp_file,ts.report_line_no - 1,3,100,(100*sz[1]/sz[0]), :fit_y=>true)
      end
    end
    
    sheet.select(1,1) # 先頭セルをセレクト
    
    # Excelを保存
    book.save
  end

# ログインの例
#
#  # ログイン用のダイアログが表示されていなければクリックして表示する。
#  if ! ts.find_element('//input[@name="session[username_or_email]"]').displayed?
#    ts.find_element('//small[text() = "アカウントをお持ちの場合"]').click
#  end
#
#  # ログイン　配列の最後のデータはサブミットボタンの指定
#  # IDとパスワードは書き換えてください。
#  #
#  ts.login([[:name,'session[username_or_email]','your_facebook_ID'],
#            [:name,'session[password]','password'],
#            [:xpath,'//input[@type="submit"]']])

  # ブラウザの終了
  ts.close
  
end
