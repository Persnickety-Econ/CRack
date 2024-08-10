# MoreScanInformationFromFilename - a script for ComicRack
# Based on ScanInformationFromFilename by Stonepaw & 600WPMPO
# New regex by the amazing Helmic
# v 0.1 (rel. 2012-March-24)


import clr, re
import System

clr.AddReference("System.Windows.Forms")

clr.AddReference("System.Drawing")

from System.IO import FileInfo

from System.Text.RegularExpressions import Regex, RegexOptions

from System.ComponentModel import BackgroundWorker

from System.Drawing import Point, Size

from System.Windows.Forms import Form, ListBox, Button, Label, TextBox, DialogResult, MessageBox, TabControl, TabPage, DockStyle, BorderStyle


#Some important constants
FOLDER = FileInfo(__file__).DirectoryName + "\\"

USERSCANPAGESFILE = FOLDER + "USERSCANPAGESFILE.txt"
USERFORMATSFILE = FOLDER + "USERFORMATSFILE.txt"
USERCOVERSFILE = FOLDER + "USERCOVERSFILE.txt"
USERFEATURESFILE = FOLDER + "USERFEATURESFILE.txt"
SETTINGSFILE = FOLDER + "settings.dat"
ICON = FOLDER + "MoreScanInformationFromFilename.ico"

BLACKLISTFILE = FOLDER + "blacklist.txt"

#@Name More Scan Information From Filename
#@Hook Books
#@Image MoreScanInformationFromFilename.png
#@Key morescaninfofromfilename

def MoreScanInformationFromFilename(books):
    
    progress = ProgressDialog(books)

    progress.ShowDialog()

    progress.Dispose()


def FindScanners(worker, books):
    
    #Load the various settings. settings is a dict
    settings = LoadSettings()

    #Load the scanners
    unformatedscanpages = LoadListFromFile(USERSCANPAGESFILE)
    unformatedcovers = LoadListFromFile(USERCOVERSFILE)
    unformatedformats = LoadListFromFile(USERFORMATSFILE)
    unformatedfeatures = LoadListFromFile(USERFEATURESFILE)

    unformatedscanpages += unformatedfeatures + unformatedformats + unformatedcovers

    #Sort the scanners by length and reverse it. For example cl will come after clickwheel allowing them to be matched correctly.
    unformatedscanpages.sort(key=len, reverse=True)

    #Format the scanners for use in the regex
    scanpages = "|".join(unformatedscanpages)
    scanpages = "(?<Tags>" + scanpages + ")"

    print scanpages
    #These amazing regex are designed by the amazing Helmic.

    blacklist = LoadListFromFile(BLACKLISTFILE)

    formatedblacklist = "|".join(blacklist)

    print formatedblacklist

    #Add in the blacklist

    #These amazing regex are designed by the amazing Helmic.

    pattern = r"(?:(?:__(?!.*__[^_]))|[(\[])(?!(?:" + formatedblacklist + r"|[\s_\-\|/,])+[)\]])(?<Tags>(?=[^()\[\]]*[^()\[\]\W\d_])[^()\[\]]{2,})[)\]]?"

    replacePattern = r"(?:[^\w]|_|^)(?:" + formatedblacklist + r")(?:[^\w]|_|$)"



    #Create the regex

    #regex = Regex(pattern, RegexOptions.IgnoreCase)
    regexScanPages = Regex(scanpages, RegexOptions.IgnoreCase)
    regexReplace = Regex(replacePattern, RegexOptions.IgnoreCase)
    
    print regexReplace


    ComicBookFields = ComicRack.App.GetComicFields()
    ComicBookFields.Remove("Scan Information")
    ComicBookFields.Add("Language", "LanguageAsText")

    for book in books:

        #.net Regex
        #Note that every possible match is found and then the last one is used.
        #This is because in some rare cases more than one thing is mistakenly matched and the scanner is almost always the last match.
        matches = regexScanPages.Matches(book.FileName)
        
      

        # try:
        #     match = matches[matches.Count-1]
            
        # except ValueError:
            
        #     #No match
        #     #print "Trying the Scanners.txt list"

        #     #Check the defined scanpages
        #     match = regexScanPages.Match(book.FileName)
        #     print match

        #     #Still no match
        #     if match.Success == False:
        #         #print "No Matches found"
        #         continue                


        # #Check if what was grabbed is field in the comic
        # fields = []
        # for field in ComicBookFields.Values:
        #     fields.append(unicode(getattr(book, field)).lower())

        # if match.Groups["Tags"].Value.lower() in fields:
        #     print "Uh oh. That matched tag is in the info somewhere."
        #     newmatch = False
        #     for n in reversed(range(0, matches.Count-1)):
        #         if not matches[n].Groups["Tags"].Value.lower() in fields:
        #             match = matches[n]
        #             newmatch = True
        #             print match
        #             break
        #     if newmatch == False:
        #         continue


        # #Check if the match can be found in () in the series, title or altseries
        # titlefields = [book.ShadowSeries, book.ShadowTitle, book.AlternateSeries]
        # abort = False
        # for title in titlefields:
        #     titleresult = re.search("\((?P<match>.*)\)", title)
        #     if titleresult != None and titleresult.group("match").lower() == match.Groups["Tags"].Value.lower():
        #         #The match is part of the title, series or altseries so skip it
        #         print "The match is part of the title, series or altseries"
        #         abort = True
        #         break
        # if abort == True:
        #     continue


        #Get a list of the old ScanInformation
        oldtags = book.ScanInformation
        ListOfTagsTemp=oldtags.split(";")
        if '' in ListOfTagsTemp:
            ListOfTagsTemp.remove('')
        
        ListOfTags=[]
        if ListOfTagsTemp != []:
            for indtag in ListOfTagsTemp:
                ListOfTags.append(indtag.strip())


        #Create our new tag
        #newtag = settings["Prefix"] + regexReplace.Replace(match.Groups["Tags"].Value.strip("_, "), "")
        
        numbers = ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen"]


        print matches.Count

        for match in matches:
            newtag =  match.Groups["Tags"].Value.strip("_, ")

            # convert newtag txt to numbers
            for j in numbers:
                newtag = newtag.replace(j,str(numbers.index(j)))

            newtag = newtag.replace("Both","2")
            newtag = newtag.replace("both","2")



            if newtag not in ListOfTags:
                ListOfTags.append(newtag)
                # print ListOfTags

        #Sort alphabeticaly to be neat
        ListOfTags.sort()


        #Add to ScanInformation field
        book.ScanInformation = "; ".join(ListOfTags)

       

#@Key morescaninfofromfilename
#@Hook ConfigScript
def MoreScanInformationFromFilenameOptions():
    
    settings = LoadSettings()

    scanpages = LoadListFromFile(USERSCANPAGESFILE)
    covers = LoadListFromFile(USERCOVERSFILE)
    formats = LoadListFromFile(USERFORMATSFILE)
    features = LoadListFromFile(USERFEATURESFILE)

    #blacklist = LoadUserBlackListFromFile(USERBLACKLISTFILE)

    optionform = OptionsForm(scanpages, covers, features, formats, settings["Prefix"])

    result = optionform.ShowDialog()

    if result == DialogResult.OK:
        settings["Prefix"] = optionform.Prefix.Text
        SaveScanPages(list(optionform.ScanPages.Items))
        SaveCovers(list(optionform.Covers.Items))
        SaveFormats(list(optionform.Formats.Items))
        SaveFeatures(list(optionform.Features.Items))
        #SaveBlackList(list(optionform.Blacklist.Items))
        SaveSettings(settings)

def LoadSettings():  ########## To fix
    #Define some default settings
    settings = {"Prefix" : "Pages:"}

    #The settings file should be formated with each line as SettingName:Value. eg Prefix:Scanner:

    try:
        with open(SETTINGSFILE, 'r') as settingsfile:
            for line in settingsfile:
                match = re.match("(?P<setting>.*?):(?P<value>.*)", line)
                settings[match.group("setting")] = match.group("value")

    except Exception, ex:
        print "Something has gone wrong loading the settings file. The error was: " + str(ex)

    return settings

def SaveSettings(settings):  ########## To fix
    
    with open(SETTINGSFILE, 'w') as settingsfile:
        for setting in settings:
            settingsfile.write(setting + ":" + settings[setting] + "\n")

def LoadListFromFile(filepath):  ########## To fix
    #The file should be formated with each list item as a new line.
    #It doesn't matter what order the scanners are in the file.
    with open(filepath, 'r') as f:
        l = f.read().splitlines()

    return l


def LoadCoversFromFile(filepath):  ########## To fix
    #The file should be formated with each list item as a new line.
    #It doesn't matter what order the scanners are in the file.
    with open(filepath, 'r') as f:
        l = f.read().splitlines()
    sl = []
    for i in l:
        sl.append(re.sub(r"[\[\]\\^$.|?*+(){}]", "", i))
    return sl

def SaveScanPages(scanpages):
    with open(USERSCANPAGESFILE, 'w') as scanpagesfile:
        for scanpage in scanpages:
            scanpagesfile.write(scanpage + "\n")

def SaveFormats(formats):
    with open(USERFORMATSFILE, 'w') as f:
        for item in formats:
            f.write(item + "\n")

def SaveFeatures(features):
    with open(USERFEATURESFILE, 'w') as f:
        for item in features:
            f.write(item + "\n")

def SaveCovers(covers):
    with open(USERCOVERSFILE, 'w') as f:
        for item in covers:
            f.write(item + "\n")

class OptionsForm(Form):
    def __init__(self, scanpages, covers, features, formats, prefixes):
        self.InitializeComponent()
        #self.Prefix.Text = prefix
        self.ScanPages.Items.AddRange(System.Array[System.String](scanpages))
        self.Covers.Items.AddRange(System.Array[System.String](covers))
        self.Features.Items.AddRange(System.Array[System.String](features))
        self.Formats.Items.AddRange(System.Array[System.String](formats))
        #self.Prefixes.Items.AddRange(System.Array[System.String](prefixes))
    
    def InitializeComponent(self):
        self.ScanPages = ListBox()
        self.Covers = ListBox()
        self.Features = ListBox()
        self.Formats = ListBox()
        self.Add = Button()
        self.Remove = Button()
        self.Prefix = TextBox()
        self.lblprefix = Label()
        self.Okay = Button()
        self.Tabs = TabControl()
        self.ScanPagesTab = TabPage()
        self.CoversTab = TabPage()
        self.FeaturesTab = TabPage()
        self.FormatsTab = TabPage()
        # 
        # ScanPages
        #
        self.ScanPages.BorderStyle = BorderStyle.None
        self.ScanPages.Dock = DockStyle.Fill
        self.ScanPages.TabIndex = 0
        self.ScanPages.Sorted = True
        # 
        # Covers
        # 
        self.Covers.Dock = DockStyle.Fill
        self.Covers.TabIndex = 0
        self.Covers.Sorted = True
        self.Covers.BorderStyle = BorderStyle.None
        # 
        # Features
        # 
        self.Features.Dock = DockStyle.Fill
        self.Features.TabIndex = 0
        self.Features.Sorted = True
        self.Features.BorderStyle = BorderStyle.None
        # 
        # Formats
        # 
        self.Formats.Dock = DockStyle.Fill
        self.Formats.TabIndex = 0
        self.Formats.Sorted = True
        self.Formats.BorderStyle = BorderStyle.None
        
        # 
        # Add
        # 
        self.Add.Location = Point(328, 102)
        self.Add.Size = Size(75, 23)
        self.Add.Text = "Add"
        self.Add.Click += self.AddItem
        # 
        # Remove
        # 
        self.Remove.Location = Point(328, 162)
        self.Remove.Size = Size(75, 23)
        self.Remove.Text = "Remove"
        self.Remove.Click += self.RemoveItem
        # 
        # Prefix
        # 
        self.Prefix.Location = Point(76, 313)
        self.Prefix.Size = Size(136, 20)
        # 
        # lblprefix
        # 
        self.lblprefix.AutoSize = True
        self.lblprefix.Location = Point(12, 36)
        self.lblprefix.Size = Size(58, 13)
        self.lblprefix.Text = "Tag Prefix:"
        # 
        # Okay
        # 
        self.Okay.Location = Point(228, 339)
        self.Okay.Size = Size(75, 23)
        self.Okay.Text = "Okay"
        self.Okay.DialogResult = DialogResult.OK
        #
        # ScanPagesTab
        #
        self.ScanPagesTab.Text = "Pagination"
        self.ScanPagesTab.UseVisualStyleBackColor = True
        self.ScanPagesTab.Controls.Add(self.ScanPages)
        #
        # CoversTab
        #
        self.CoversTab.Text = "Covers"
        self.CoversTab.UseVisualStyleBackColor = True
        self.CoversTab.Controls.Add(self.Covers)
        #
        # FeaturesTab
        #
        self.FeaturesTab.Text = "Scan Features"
        self.FeaturesTab.UseVisualStyleBackColor = True
        self.FeaturesTab.Controls.Add(self.Features)
        #
        # FormatsTab
        #
        self.FormatsTab.Text = "Comic Format"
        self.FormatsTab.UseVisualStyleBackColor = True
        self.FormatsTab.Controls.Add(self.Formats)
        #
        # Tabs
        #
        self.Tabs.Size = Size(310, 280)
        self.Tabs.Location = Point(12, 12)
        self.Tabs.Controls.Add(self.ScanPagesTab)
        self.Tabs.Controls.Add(self.CoversTab)
        self.Tabs.Controls.Add(self.FormatsTab)
        self.Tabs.Controls.Add(self.FeaturesTab)
        # 
        # Form Settings
        # 
        self.Size = System.Drawing.Size(450, 400)
        self.Controls.Add(self.Tabs)
        self.Controls.Add(self.Add)
        self.Controls.Add(self.Remove)
        self.Controls.Add(self.lblprefix)
        self.Controls.Add(self.Prefix)
        self.Controls.Add(self.Okay)
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        self.Text = "MORE Scan Information From Filename Options"
        self.MinimizeBox = False
        self.MaximizeBox = False
        self.AcceptButton = self.Okay
        self.Icon = System.Drawing.Icon(ICON)


    def RemoveItem(self, sender, e):
        if self.Tabs.SelectedTab == self.ScanPagesTab:
            self.ScanPages.Items.Remove(self.ScanPages.SelectedItem)
        elif self.Tabs.SelectedTab == self.CoversTab:
            self.Covers.Items.Remove(self.Covers.SelectedItem)
        elif self.Tabs.SelectedTab == self.FeaturesTab:
            self.Features.Items.Remove(self.Features.SelectedItem)
        else:
            self.Formats.Items.Remove(self.Formats.SelectedItem)

    def AddItem(self, sender, e):
        input = InputBox()
        input.Owner = self
        if input.ShowDialog() == DialogResult.OK:
            if self.Tabs.SelectedTab == self.ScanPagesTab:
                self.ScanPages.Items.Add(input.FindName())
                self.ScanPages.SelectedItem = input.FindName()
            elif self.Tabs.SelectedTab == self.CoversTab:
                self.Covers.Items.Add(input.FindName())
                self.Covers.SelectedItem = input.FindName()
            elif self.Tabs.SelectedTab == self.FeaturesTab:
                self.Features.Items.Add(input.FindName())
                self.Features.SelectedItem = input.FindName()
            else:
                self.Formats.Items.Add(input.FindName())
                self.Formats.SelectedItem = input.FindName()                

class InputBox(Form):
    def __init__(self):
        self.TextBox = TextBox()
        self.TextBox.Size = Size(250, 20)
        self.TextBox.Location = Point(15, 12)
        self.TextBox.TabIndex = 1
        
        self.OK = Button()
        self.OK.Text = "OK"
        self.OK.Size = Size(75, 23)
        self.OK.Location = Point(109, 38)
        self.OK.DialogResult = DialogResult.OK
        self.OK.Click += self.CheckTextBox
        
        self.Cancel = Button()
        self.Cancel.Size = Size(75, 23)
        self.Cancel.Text = "Cancel"
        self.Cancel.Location = Point(190, 38)
        self.Cancel.DialogResult = DialogResult.Cancel
        
        self.Size = Size(300, 100)
        self.Text = "Please enter a new search expression"
        self.Controls.Add(self.OK)
        self.Controls.Add(self.Cancel)
        self.Controls.Add(self.TextBox)
        self.AcceptButton = self.OK
        self.CancelButton = self.Cancel
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        self.Icon = System.Drawing.Icon(ICON)
        self.ActiveControl = self.TextBox
        
    def FindName(self):
        if self.DialogResult == DialogResult.OK:
            return self.TextBox.Text.strip()
        else:
            return None
        
    def CheckTextBox(self, sender, e):
        if not self.TextBox.Text.strip():
            MessageBox.Show("Please enter a name into the textbox")
            self.DialogResult = DialogResult.None
        
        if self.TextBox.Text.strip() in self.Owner.ScanPages.Items:
            MessageBox.Show("The entered name is already in the Pages tab. Please enter another")
            self.DialogResult = DialogResult.None
        
        if self.TextBox.Text.strip() in self.Owner.Formats.Items:
            MessageBox.Show("The entered name is already in the Formats tab. Please enter another")
            self.DialogResult = DialogResult.None

        if self.TextBox.Text.strip() in self.Owner.Features.Items:
            MessageBox.Show("The entered name is already in the Features tab. Please enter another")
            self.DialogResult = DialogResult.None

        if self.TextBox.Text.strip() in self.Owner.Covers.Items:
            MessageBox.Show("The entered name is already in the Covers tab. Please enter another")
            self.DialogResult = DialogResult.None

        # if self.TextBox.Text.strip() in self.Owner.Formats.Items:
        #     MessageBox.Show("The entered name is already in the Formats tab. Please enter another")
        #     self.DialogResult = DialogResult.None

        # if self.TextBox.Text.strip() in self.Owner.Features.Items:
        #     MessageBox.Show("The entered name is already in the Features tab. Please enter another")
        #     self.DialogResult = DialogResult.None

        # if self.TextBox.Text.strip() in self.Owner.Covers.Items:
        #     MessageBox.Show("The entered name is already in the Covers tab. Please enter another")
        #     self.DialogResult = DialogResult.None

class ProgressDialog(Form):
    
    def __init__(self, books):
        self.InitializeComponent()
        self.worker.RunWorkerAsync(books)
        self.done = False

    def InitializeComponent(self):
        self.progressBar = System.Windows.Forms.ProgressBar()
        # 
        # progressBar
        # 
        self.progressBar.Location = Point(0, 0)
        self.progressBar.Size = Size(284, 23)
        self.progressBar.Maximum = 100
        self.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee

        self.worker = BackgroundWorker()
        self.worker.DoWork += self.WorkerDoWork
        self.worker.RunWorkerCompleted += self.WorkerCompleted

        self.ClientSize = Size(284, 23)
        self.Controls.Add(self.progressBar)
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.Text = "Searching Filenames"
        self.Icon = System.Drawing.Icon(ICON)
        self.FormClosing += self.CheckClosing
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen

    #Stop the user from closing the progress dialog
    def CheckClosing(self, sender, e):
        if e.CloseReason == System.Windows.Forms.CloseReason.UserClosing and self.done == False:
            e.Cancel = True

    def SetTitle(self, text):
        self.Text = text

    def WorkerDoWork(self, sender, e):
        FindScanners(sender, e.Argument)
        e.Result = "Done"

    def WorkerCompleted(self, sender, e):
        self.done = True
        self.Close()
