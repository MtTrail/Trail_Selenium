#! ruby -EWindows-31J
# -*- mode:ruby; coding:Windows-31J -*-

#= Trail4You Selenium�g�����W���[��
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

# SSL�ؖ����ŃG���[���N����̂�h���ׂɖ�������
# �ؖ����ɖ�肪������΃R�����g�A�E�g���Ă�������
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
#==�ړI
# Selenium�Ńf�[�^���W���邽�߂̃N���X
#*   Excel�N���X�����p����B
#*   Windows�p�ł��B
#*   �p�X��'/'��؂�ň����܂��B'\\'�ł͗L��܂���B
#
class Trail_Selenium
  attr_accessor :driver,:wait
  attr_accessor :report_book,:report_sheet,:report_line_no
  
  #=== ������ wait����(�b)���w�肷��B
  def initialize (wait_time = 10)
    @driver = Selenium::WebDriver.for :firefox
    @wait = Selenium::WebDriver::Wait.new(:timeout => wait_time) # seconds
    @report_book = nil
    @report_sheet = nil
    @report_line_no = 1
  end

  #=== ���O�C��
  #    �����Őݒ�l�̔z��̔z���n���A��ԍŌ��submit�{�^���̏��(�ݒ�l�Ȃ�)
  #    �e�z��v�f�͉��L�̌`��
  #    [�������V���{��,�����̒l,�ݒ�l]�@���́@[:xpath, 'xpath�w��',�ݒ�l]
  #    �� : login([[:name,'UserName','LoginName'],[:name,'Password','password'],[:name,'Submit']])
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
  
  #=== �f�[�^�������ݗp��Excel���w��
  #  target : open����Excel�t�@�C���̃p�X
  #  tenplate : �e���v���[�g��excel�t�@�C���̃p�X�A�����target�ɃR�s�[���Ă���open����B
  #
  #   �e���v���[�g���w�肷��Ƃ�����R�s�[���Ďg�p����B
  #   �u���b�N�ŏ������e���󂯎��
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

  #=== Excel�V�[�g�̃I�[�v��
  #  book : open���ꂽexcel�I�u�W�F�N�g
  #  sheet_name : �V�[�g�ԍ��܂��̓V�[�g��
  #
  #   �u�b�N�ƃV�[�g�����w�肷��B
  #   �V�[�g�������l�̏ꍇ�V�[�g�̔ԍ��ƌ��Ȃ����
  #   Excel�̃V�[�g�I�u�W�F�N�g��Ԃ�
  #
  #   ��
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

  #=== ���|�[�g�p��Excel�V�[�g�̃I�[�v��
  #  book : open���ꂽexcel�I�u�W�F�N�g
  #  sheet_name : �V�[�g�ԍ��܂��̓V�[�g��
  #
  # @report_sheet��ݒ肷��
  # Excel�ɏ������ޏꍇ�A��������w�肷��Ə������݊֐��Ăяo�����Ƀp�����[�^�����点��B
  #
  def open_report_sheet( book, sheet_name )
    @report_sheet = open_sheet( book, sheet_name )
    @report_sheet
  end


  #-------------------------------------------------------------------

  #=== �R���\�[���ւ̕\����Report�V�[�g�ւ̋L�^
  #   offset: �̓R���\�[���o�͎��̍��}�[�W���Ƃ��Ďg�p�����B
  #         : �܂�Excel�V�[�g�̏ꍇ�A���J�����ڂ���f�[�^���Z�b�g���邩�̎w��ƂȂ�B
  #   t     : �o�͂��镶����̔z����w�肷��B�R���\�[����Excel�V�[�g�ɏo�͂����B
  #   sheet : �f�t�H���g��@report_sheet�ȊO�̃V�[�g�ɏo�͂���Ƃ��n�b�V���Ŏw�肷��B:sheet => other_sheet
  #   line_no : �o�͂̍s�ԍ����f�t�H���g��@report_line_no�ȊO�̂Ƃ��n�b�V���Ŏw�肷��B :line_no => 2
  #
  #   �o�͂���Excel�V�[�g�Əo�͍s�̓n�b�V���Ŏw�肷��B�w�肳��Ȃ��ꍇ�A�Ō��open_report_sheet�ŊJ�����V�[�g���g����B
  #   �o�͂���s�͎w�肳��Ȃ��ꍇ @report_line_no���g�p�����B
  #
  #   �o�͌��sheet���w�肳��Ă��Ȃ��ꍇ @report_line_no�� + 1 �����B
  #   �o�͕�����ɃJ���}�����܂܂Ȃ��Ƃ������������͂��邪offset=0�̃R���\�[���o�͂��t�@�C���Ƀ��_�C���N�g�����CSV�t�@�C���ƂȂ�B
  #
  #   ���� : Excel�֏o�͂���ꍇopen_excel�̃u���b�N���ŗ��p����Ȃ���΂Ȃ�Ȃ��B
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

  #=== xpath�Ŏw�肳�ꂽ�摜�G�������g����摜���R�s�[�@�\���g�p���A�N���b�v�{�[�h�o�R�Ńt�@�C���ɗ��Ƃ��B
  # ���I�ɐ��������摜��ۑ�����Ƃ��Ɏg�p����B
  # �E�N���b�N�ŉ摜���R�s�[���j���[���o�Ȃ����̂ɂ͎g�p�ł��Ȃ��B
  #
  #   node     : xpath�̊J�n�m�[�h
  #   xp       : �摜���w�肷��xpath
  #   filename : �������މ摜�t�@�C����
  #   wait_mode: �G�������g�̏o����҂Ƃ� true ���w�� :wait_mode=>true
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

  #=== xpath�Ŏw�肳�ꂽ�摜�G�������g(img�^�O)��URL����摜���t�@�C���ɗ��Ƃ��B
  #   node     : xpath�̊J�n�m�[�h
  #   xp       : �摜���w�肷��xpath
  #   pathname : �摜�t�@�C�����������ރt�H���_�p�X
  #   wait_mode: �G�������g�̏o����҂Ƃ� true ���w�� :wait_mode=>true
  #   rename   : �t�@�C���������̖��O���珑��������Ƃ��w�� :rename => 'newname.jpg'
  #            : �w�肳��Ȃ����src�����Ɏw�肳�ꂽ�t�@�C�������g�p�����B
  #
  #   <return> : �摜�t�@�C���p�X or nil
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

  #=== Excel�ɉ摜�\��t��
  #  filename : �\��t����摜�t�@�C��(Excel���Ɏ�荞�܂��)
  #  cx,cy : �\��t���ʒu�̃J����(cx)�ƍs(cy) 1�n�܂�̒l
  #  sh,sw : �摜�̓\��t���h�b�g�T�C�Y ����(sh) ��(sw)
  #  sheet : �\��t����V�[�g�I�u�W�F�N�g�A�w�薳���̏ꍇ@report_sheet
  #  fit_x : �J��������sw�ɍ��킹��B
  #  fit_y : �s�̍�����sh�ɍ��킹��B
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


  #=== �Z���N�gBOX�I��
  #   xp : �Z���N�g�G�������g���w�肷��xpath
  #   tx : �I�����镶����̓��e
  #
  def select_by_text(xp,tx)
    select = Selenium::WebDriver::Support::Select.new( @wait.until{@driver.find_element(:xpath,xp)} )
    select.select_by(:text,tx.encode('UTF-8'))
  end
  
  
  #=== �G�������g�T��
  #  xp   : �Z���N�g�G�������g���w�肷��xpath
  #  node : �r���̗v�f����̏ꍇ�A���̗v�f�I�u�W�F�N�g���w�肷�� :node => element
  #
  #   ������Ȃ��Ƃ��ɂ�nil��Ԃ��B
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

  #=== �G�������g�T��(������܂ő҂�)
  #  xp   : �Z���N�g�G�������g���w�肷��xpath
  #  node : �r���̗v�f����̏ꍇ�A���̗v�f�I�u�W�F�N�g���w�肷�� :node => element
  #
  #   ������Ȃ��Ƃ��ɂ�nil��Ԃ��B
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
  
  #=== Selenium �I��
  #
  def close
    @driver.quit()
  end

end   ## --- end of Trail_Selenium class ---


if $0 == __FILE__

require 'date'

  #
  # Twitter�̃��b�Z�[�W�Ɖ摜������ł��B
  #
  ts = Trail_Selenium.new    # �N���X�̏�����
  driver = ts.driver
  driver.navigate.to( 'https://twitter.com/3190')  # �䍂�R������̃y�[�W

  # ���O�C���̃_�C�A���O���o�Ă��������
  if ts.find_element('//input[@name="session[username_or_email]"]').displayed?
    ts.find_element('//small[text() = "�A�J�E���g���������̏ꍇ"]').click
  end

  # �ŏ��̉摜�� pic1.jpg �Ƃ������O�ŕۑ����Ă݂�B
  ts.get_picture('','//div[@class="AdaptiveMedia-photoContainer js-adaptive-photo "]/img',:rename=>'Hotaka.jpg')
  
  # ���ݓǂݍ��܂�Ă���post�����o��
  posts = driver.find_elements(:xpath,'//div[@class="stream"]/ol/li')
  print "Post�� = #{posts.size}\n"
  
  # �ۑ��p��Excel�t�@�C�����e���v���[�g����R�s�[���ĊJ��
  ts.open_excel("Hotakasansou_#{DateTime.now.strftime("%Y%m%d%H%M%S")}.xls",'Report_tenplate.xls') do |book|

    # �������ݗp�̃V�[�g���J��
    sheet = ts.open_report_sheet(book,'Report')

    ts.report_line_no = 2      # �����݊J�n��2�s�ڂ���
    sheet.set_width(1,1,11.8)  # �Z��(1,1)��100pixcel�̃T�C�Y�ɂ���B
    sheet.set_height(1,1,75)   # Y: 100pixcel = 75 point
    sheet.set_width(1,3,11.8)  # X: 100pixcel = 11.8 ����  �摜�p�̃Z��������ݒ�

    # �ǂݍ���post�����ɏ���
    posts.each do |post|
      
      # �c�C�[�g�����l�̖��O
      name = post.find_element(:xpath,'.//div[@class="stream-item-header"]/a/strong').text()
      
      # ���b�Z�[�W���e
      msg  = post.find_element(:xpath,'.//p').text()
      
      # �R���\�[����Excel�֏o��
      ts.disp_msg_array(0,[name,msg])

      # ��΃p�X�ł̃t�@�C�������쐬
      temp_file = File.expand_path('../temp.jpg',__FILE__)

      # �摜������Ȃ���o��
      if ts.get_picture_via_clipboard(temp_file, './/div[@class="content"]/div[3]//img',:node=>post)
        # �s�N�Z���T�C�Y�𒲂ׂ�
        sz = FastImage.size(temp_file)
        
        # Excel�ɓ\��t��report_line_no�͊��Ɏ��̍s�ɍs���Ă���̂ŕ␳����
        # �摜�T�C�Y�͕���100pixcel�ɂȂ�悤�ɕ␳
        # fit_y��true�ɂ��ăZ���̍������摜�̍����ɍ��킹��B
        ts.add_picture_to_excel(temp_file,ts.report_line_no - 1,3,100,(100*sz[1]/sz[0]), :fit_y=>true)
      end
    end
    
    sheet.select(1,1) # �擪�Z�����Z���N�g
    
    # Excel��ۑ�
    book.save
  end

# ���O�C���̗�
#
#  # ���O�C���p�̃_�C�A���O���\������Ă��Ȃ���΃N���b�N���ĕ\������B
#  if ! ts.find_element('//input[@name="session[username_or_email]"]').displayed?
#    ts.find_element('//small[text() = "�A�J�E���g���������̏ꍇ"]').click
#  end
#
#  # ���O�C���@�z��̍Ō�̃f�[�^�̓T�u�~�b�g�{�^���̎w��
#  # ID�ƃp�X���[�h�͏��������Ă��������B
#  #
#  ts.login([[:name,'session[username_or_email]','your_facebook_ID'],
#            [:name,'session[password]','password'],
#            [:xpath,'//input[@type="submit"]']])

  # �u���E�U�̏I��
  ts.close
  
end
