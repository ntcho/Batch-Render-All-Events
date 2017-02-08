/**
 * Sample script that performs batch renders with GUI for selecting
 * render templates.
 *
 * Revision Date: Jun. 28, 2006.
 *
 *
 *
 *  ----------= BatchRender+Events =--------------
 *
 * Название: BatchRender+Events.cs
 *    Автор: Дмитрий ЧайнеГ <chainick@narod.ru>
 *     Дата: 05.08.2012
 *   Версия: 0.0.9
 * Описание: Скрипт "BatchRender+Events.cs" полностью
 *           основан на "Batch Render.cs" из официальной
 *           версии Sony Vegas 10. Функционал оригинального
 *           скрипта полностью сохранен. "BatchRender+Events.cs"
 *           позволяет:
 *           - перекодировать каждый фрагмент выделенного
 *             трека в отдельный файл с его размещением
 *             на новом треке;
 *           - создать offline-фрагменты на новом треке и bat-файл
 *             для внешнего рендеринга.
 **/


using System;
using System.IO;
using System.Text;
using System.Drawing;
using System.Collections;
using System.Diagnostics;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using Sony.Vegas;


public class EntryPoint
{

    /* ---------------- Настройки скрипта ---------------- */

    // заменить существующие файлы.
    //   true  - заменить
    //   false - не заменять (появится предупреждение)
    bool OverwriteExistingFiles = false;

    // путь по умолчанию
    string myDefBasePath = "";    // @"D:\render\";

    // пресет рендера по умолчанию при старте скрипта
    string myDefRenderTempl = ""; // "Avid DNxHD 1920x1080-25p 185Mbps";

    // количество кадров захлестов по умолчанию
    int myDefMarginFrames = 0;    // 25;

    // Нумеровать файлы, начиная с
    int myDefStartFileNumber = 1;

    // формат (количество знаков) цифрового индекса файла,
    // например, 00001.mov
    string myDefFormatFileNumber = "00000";

    // имя bat-файла по умолчанию
    string DefBatFileName = "BatchEncodeEvents.bat";

    // имя файла с пресетами для внешнего рендеринга
    string FilePresetList = "PresetList.txt";

    /* -------------------------------------------------------- */


    // список команд (см. ниже) для внешнего рендеринга
    Dictionary<string, string> ExtRenderCommand = new Dictionary<string, string>();

    // заголовок скрипта
    string ScriptCaption = "Batch Render + Events";

    // текущая версия скрипта
    string ScriptVersion = "0.0.9";

    string defaultBasePath = "Untitled_";

    Sony.Vegas.Vegas myVegas = null;

    enum RenderMode
    {
        Project = 0,
        Selection,
        Regions,
        Events,
        OfflineEvents,
    }

    ArrayList SelectedTemplates = new ArrayList();

    public void FromVegas(Vegas vegas)
    {
        myVegas = vegas;

        string projectPath = myVegas.Project.FilePath;

        if (false == String.IsNullOrEmpty(myDefBasePath))
        {
            string dir = Path.GetDirectoryName(myDefBasePath);
            string fileName = Path.GetFileNameWithoutExtension(myDefBasePath);
            defaultBasePath = Path.Combine(dir, fileName);

            if (String.IsNullOrEmpty(fileName))
                defaultBasePath += Path.DirectorySeparatorChar;

        }
        else if (String.IsNullOrEmpty(projectPath))
        {
            string dir = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);
            defaultBasePath = Path.Combine(dir, defaultBasePath);
        }
        else
        {
            string dir = Path.GetDirectoryName(projectPath);
            string fileName = Path.GetFileNameWithoutExtension(projectPath);
            defaultBasePath = Path.Combine(dir, fileName + "_");
        }

        DialogResult result = ShowBatchRenderDialog();
        myVegas.UpdateUI();
        if (DialogResult.OK == result)
        {
            // inform the user of some special failure cases
            string outputFilePath = FileNameBox.Text;
            RenderMode renderMode = RenderMode.Project;
            if (RenderRegionsButton.Checked)
            {
                renderMode = RenderMode.Regions;
            }
            else if (RenderSelectionButton.Checked)
            {
                renderMode = RenderMode.Selection;
            }
            else if (RenderEventsButton.Checked)
            {
                renderMode = RenderMode.Events;
            }
            else if (OfflineEventsButton.Checked)
            {
                renderMode = RenderMode.OfflineEvents;
            }

            DoBatchRender(SelectedTemplates, outputFilePath, renderMode);
        }
    }

    void DoBatchRender(ArrayList selectedTemplates, string basePath, RenderMode renderMode)
    {
        string outputDirectory = Path.GetDirectoryName(basePath);
        string baseFileName = Path.GetFileName(basePath);

        // make sure templates are selected
        if ((null == selectedTemplates) || (0 == selectedTemplates.Count))
            throw new ApplicationException("No render templates selected.");

        // make sure the output directory exists
        if (!Directory.Exists(outputDirectory))
            throw new ApplicationException("The output directory does not exist.");

        RenderStatus status = RenderStatus.Canceled;


        // формат (количество знаков) цифрового индекса файла
        string TempStr = StartFileNumberBox.Text.Trim();
        int TempInt = TempStr.Length > 0 ? TempStr.Length : myDefFormatFileNumber.Length;
        string FormatFileNumber = new String('0', TempInt);

        // начальный номер (имя) файла
        int EventFileNumber = StringToInt(StartFileNumberBox.Text, myDefStartFileNumber);

        // количество кадров захлестов
        int MarginFrames = StringToInt(MarginFramesBox.Text, myDefMarginFrames);

        // получить выделенные видеотреки
        ArrayList SelectedTracks = GetSelectedTrack(MediaType.Video);


        /* --- команды для внешнего рендеринга --- */

        // Дополснительные пресеты описываюся в файле
        // из переменной FilePresetList.

        // пресеты доступные по умолчанию.
        ExtRenderCommand.Add("All Files (*.*)|*.*", "{0} | {1} | {2} | {3}");

        // FFmpeg: ProRes
        ExtRenderCommand.Add("FFmpeg: ProRes [36Mbps 10-bit, Audio Copy] (*.bat)|*.bat",
            "ffmpeg -i \"{0}\" -ss {2} -t {3} -vcodec prores -profile 0 -acodec copy \"{1}\"");
        ExtRenderCommand.Add("FFmpeg: ProRes [75Mbps 10-bit, Audio Copy] (*.bat)|*.bat",
            "ffmpeg -i \"{0}\" -ss {2} -t {3} -vcodec prores -profile 1 -acodec copy \"{1}\"");
        ExtRenderCommand.Add("FFmpeg: ProRes [112Mbps 10-bit, Audio Copy] (*.bat)|*.bat",
            "ffmpeg -i \"{0}\" -ss {2} -t {3} -vcodec prores -profile 2 -acodec copy \"{1}\"");
        ExtRenderCommand.Add("FFmpeg: ProRes [185Mbps 10-bit, Audio Copy] (*.bat)|*.bat",
            "ffmpeg -i \"{0}\" -ss {2} -t {3} -vcodec prores -profile 3 -acodec copy \"{1}\"");

        // загрузка дополнительных пресетов из файла
        ExtRenderCommand = PresetsFromFile(ExtRenderCommand, FilePresetList);


        // enumerate through each selected render template
        foreach (RenderItem renderItem in selectedTemplates)
        {
            // construct the file name (most of it)
            string filename = Path.Combine(outputDirectory,
                FixFileName(baseFileName) +
                FixFileName(renderItem.Renderer.FileTypeName) +
                "_" +
                FixFileName(renderItem.Template.Name));

            if (RenderMode.Regions == renderMode)
            {
                int regionIndex = 0;
                foreach (Sony.Vegas.Region region in myVegas.Project.Regions)
                {
                    string regionFilename = String.Format("{0}[{1}]{2}",
                        filename,
                        regionIndex.ToString(),
                        renderItem.Extension);

                    // Render the region
                    status = DoRender(regionFilename, renderItem, region.Position, region.Length);
                    if (RenderStatus.Canceled == status)
                        break;

                    regionIndex++;
                }
            }

            /* ---------- + Events --------------  */

            else if (RenderMode.OfflineEvents == renderMode)
            {

                // если выделенные треки отсутствуют
                if (0 == SelectedTracks.Count)
                    throw new ApplicationException("No tracks selected.");

                // если фрагменты на выделенных треках отсутствуют
                if (0 == GetEventsCount(SelectedTracks))
                    throw new ApplicationException("No events on selected tracks.");

                // имя bat-файла
                // string BatFileName = DefBatFileName;

                // выбранных пресетов несколько
                // if (selectedTemplates.Count > 1)
                // {
                    // добавить в имя bat-файла название пресета
                    StringBuilder fileName = new StringBuilder("");
                    fileName.Append(Path.GetFileNameWithoutExtension(DefBatFileName));
                    fileName.AppendFormat("_[{0}]", renderItem.Template.Name);
                    fileName.Append(Path.GetExtension(DefBatFileName));
                    string BatFileName = FixFileName(fileName.ToString());
                // }

                // Диалог сохранения bat-файла
                SaveFileDialog saveFileDialog = SaveBatchFileDialog(ExtRenderCommand, BatFileName);

                // Если кнопка "Сохранить" не нажата
                if (System.Windows.Forms.DialogResult.OK != saveFileDialog.ShowDialog())
                    continue;

                // массив параметров кодирования
                ArrayList ListOptions = new ArrayList();

                // обход выделенных треков
                foreach (Track CurrentTrack in SelectedTracks)
                {
                    // если на текущем треке отсутствуют фрагменты
                    if (0 == CurrentTrack.Events.Count)
                        continue;

                    // имя целевого трека
                    StringBuilder TargetTrackName = new StringBuilder("");
                    TargetTrackName.Append(CurrentTrack.Name);
                    TargetTrackName.AppendFormat(" [ {0} ]", renderItem.Template.Name);

                    // создать (над текущим треком) целевой трек
                    Track TargetTrack = new VideoTrack(CurrentTrack.Index, TargetTrackName.ToString());
                    myVegas.Project.Tracks.Add(TargetTrack);

                    // отключить видеотрек
                    TargetTrack.Mute = true;

                    // обход фрагментов текущего трека
                    foreach (TrackEvent CurrentEvent in CurrentTrack.Events)
                    {
                        // полный путь к создаваемому файлу
                        string TargetEventFilePath = Path.Combine(outputDirectory,
                            FixFileName(baseFileName) +
                            EventFileNumber.ToString(FormatFileNumber) +
                            renderItem.Extension);

                        // статус рендеринга
                        status = RenderStatus.Complete;

                        // создание фрагмента
                        VideoEvent TargetEvent = new VideoEvent(CurrentEvent.Start, CurrentEvent.Length);

                        // добавление фрагмента на трек
                        TargetTrack.Events.Add(TargetEvent);

                        // копирование затуханий и переходов
                        CopyFades(CurrentEvent, TargetEvent);

                        // добавление медиа в проект
                        Media TargetMedia = new Media(TargetEventFilePath);

                        MediaStream TargetStream = TargetMedia.CreateOfflineStream(MediaType.Video);
                        Take TargetTake = new Take(TargetStream);
                        TargetEvent.Takes.Add(TargetTake);

                        // копирование некоторых атрибутов из исходного медиа
                        TargetMedia.TapeName = CurrentEvent.ActiveTake.Media.TapeName;
                        TargetMedia.Comment = CurrentEvent.ActiveTake.Media.Comment;

                        // захлесты справа и слева
                        Timecode MarginLeft = Timecode.FromFrames(MarginFrames);
                        Timecode MarginRight = Timecode.FromFrames(MarginFrames);

                        // проверка наличия кадров слева
                        if (CurrentEvent.ActiveTake.Offset < MarginLeft)
                            MarginLeft = CurrentEvent.ActiveTake.Offset;

                        // проверка наличия кадров справа
                        Timecode CurrentMediaLength = CurrentEvent.ActiveTake.Media.Length;
                        CurrentMediaLength = CurrentMediaLength - TargetEvent.ActiveTake.Offset - TargetEvent.Length - MarginLeft;
                        if (CurrentMediaLength < MarginRight)
                            MarginRight = CurrentMediaLength;


                        // смещение (на момент написания не работало)
                        TargetEvent.ActiveTake.Offset = CurrentEvent.ActiveTake.Offset - MarginLeft;

                        // параметры кодирования
                        Timecode ss = CurrentEvent.ActiveTake.Offset - MarginLeft;
                        Timecode t = TargetEvent.Length + MarginLeft + MarginRight;

                        // {0} - полный путь к исходному файлу
                        string OptInputFile = CurrentEvent.ActiveTake.MediaPath;
                        // {1} - полный путь к создаваемому файлу
                        string OptOutputFile = TargetEventFilePath;
                        // {2} - временная отметка начала рендеринга
                        string OptOffset = ss.ToString(RulerFormat.Time);
                        // {3} - длительность
                        string OptDuration = t.ToString(RulerFormat.Time);
                        // {4} - резерв для доп. опций (пока не реализовано)
                        //

                        // массив параметров
                        ListOptions.Add(OptInputFile + "|" + OptOutputFile + "|" + OptOffset + "|" + OptDuration);

                        EventFileNumber++;

                    } // обход фрагментов текущего трека

                    // Список вариантов рендеринга (ключи словаря)
                    List<string> KeysRenders = new List<string>(ExtRenderCommand.Keys);

                    // Плейсхолдер команды
                    string Placeholder = ExtRenderCommand[KeysRenders[saveFileDialog.FilterIndex - 1]];

                    // генерация содержимого bat-файла
                    string FileContent = BatFileContent(Placeholder, ListOptions);

                    // Сохранение файла
                    FileStream fs = File.Open(saveFileDialog.FileName, FileMode.Create, FileAccess.Write);
                    StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8);
                    sw.WriteLine(FileContent);
                    sw.Close();

                } // обход выделенных треков

            }
            else if (RenderMode.Events == renderMode)
            {

                // если выделенные треки отсутствуют
                if (0 == SelectedTracks.Count)
                    throw new ApplicationException("No tracks selected.");

                // если фрагменты на выделенных треках отсутствуют
                if (0 == GetEventsCount(SelectedTracks))
                    throw new ApplicationException("No events on selected tracks.");

                // обход выделенных треков
                foreach (Track CurrentTrack in SelectedTracks)
                {

                    // если на текущем треке отсутствуют фрагменты
                    if (0 == CurrentTrack.Events.Count)
                        continue;

                    // имя целевого трека
                    StringBuilder TargetTrackName = new StringBuilder("");
                    TargetTrackName.Append(CurrentTrack.Name);
                    TargetTrackName.AppendFormat(" [ {0} ]", renderItem.Template.Name);

                    // создать (над текущим треком) целевой трек
                    Track TargetTrack = new VideoTrack(CurrentTrack.Index, TargetTrackName.ToString());
                    myVegas.Project.Tracks.Add(TargetTrack);

                    // отключить видеотрек
                    TargetTrack.Mute = true;

                    // обход фрагментов текущего трека
                    foreach (TrackEvent CurrentEvent in CurrentTrack.Events)
                    {

                        // отключить фрагменты за исключением текущего
                        EventsMuteOn (CurrentTrack, CurrentEvent);

                        // оригинальные параметры текущего фрагмента
                        Timecode CurrentEventLength = CurrentEvent.Length;
                        Timecode CurrentEventStart = CurrentEvent.Start;
                        Timecode CurrentEventActiveTakeOffset = CurrentEvent.ActiveTake.Offset;
                        Timecode CurrentEventFadeInLength = CurrentEvent.FadeIn.Length;
                        Timecode CurrentEventFadeOutLength = CurrentEvent.FadeOut.Length;
                        
                        // захлесты справа и слева
                        Timecode MarginLeft = Timecode.FromFrames(MarginFrames);
                        Timecode MarginRight = Timecode.FromFrames(MarginFrames);

                        // проверка наличия кадров слева
                        if (CurrentEvent.ActiveTake.Offset < MarginLeft)
                            MarginLeft = CurrentEvent.ActiveTake.Offset;

                        // проверка наличия кадров справа
                        Timecode CurrentMediaLength = CurrentEvent.ActiveTake.Media.Length;
                        CurrentMediaLength = CurrentMediaLength - CurrentEvent.ActiveTake.Offset - CurrentEvent.Length - MarginLeft;
                        if (CurrentMediaLength < MarginRight)
                            MarginRight = CurrentMediaLength;


                        // временные параметры текущего фрагмента (для рендеринга):
                        CurrentEvent.FadeIn.Length = new Timecode(0);
                        CurrentEvent.FadeOut.Length = new Timecode(0);
                        CurrentEvent.Start =  CurrentEvent.Start - MarginLeft;
                        CurrentEvent.ActiveTake.Offset = CurrentEvent.ActiveTake.Offset - MarginLeft;
                        CurrentEvent.Length = CurrentEvent.Length + MarginLeft + MarginRight;
						

						// зменение параметров соседних фрагментов:
						// осуществлено из-за влияния наличия переходов 
						// на рендеринг текущего фрагмента
						TrackEvent PrevEvent = null;
						TrackEvent NextEvent = null;
						Timecode PrevEventEnd = null;
						Timecode NextEventStart = null;

						if (0 != CurrentEvent.Index)
						{
							PrevEvent = CurrentTrack.Events[CurrentEvent.Index - 1];
							PrevEventEnd = PrevEvent.End;
							PrevEvent.End = CurrentEvent.Start;
						}

						if ((CurrentTrack.Events.Count - 1) != CurrentEvent.Index)
						{
							NextEvent = CurrentTrack.Events[CurrentEvent.Index + 1];
							NextEventStart = NextEvent.Start;
							NextEvent.Start = CurrentEvent.End;
						}


                        /* --- Рендеринг фрагмента в отдельный файл -- */

                        // полный путь к создаваемому файлу
                        string TargetEventFilePath = Path.Combine(outputDirectory,
                             FixFileName(baseFileName) +
                             EventFileNumber.ToString(FormatFileNumber) +
                             renderItem.Extension);

                        // рендеринг
                        status = DoRender(TargetEventFilePath, renderItem, CurrentEvent.Start, CurrentEvent.Length);
                        status = RenderStatus.Complete;


						// восстановление параметров соседних фрагментов
						if (null != PrevEvent)
							PrevEvent.End = PrevEventEnd;
						
						if (null != NextEvent)
							NextEvent.Start = NextEventStart;

                        // восстановление параметров текущего фрагмента
                        CurrentEvent.Length = CurrentEventLength;
                        CurrentEvent.Start = CurrentEventStart;
                        CurrentEvent.ActiveTake.Offset = CurrentEventActiveTakeOffset;
                        CurrentEvent.FadeIn.Length = CurrentEventFadeInLength;
                        CurrentEvent.FadeOut.Length = CurrentEventFadeOutLength;


                        /* --- размещение нового фрагмента файла на конечном видеотреке --- */

                        // добавление файла в проект
                        Media TargetMedia = new Media(TargetEventFilePath);

                        // создание фрагмента
                        VideoEvent TargetEvent = new VideoEvent(CurrentEvent.Start, CurrentEvent.Length);

                        // добавление фрагмента на трек
                        TargetTrack.Events.Add(TargetEvent);

                        // копирование затуханий и переходов
                        CopyFades(CurrentEvent, TargetEvent);

                        MediaStream TargetStream = TargetMedia.Streams.GetItemByMediaType(MediaType.Video, 0);
                        Take TargetTake = new Take(TargetStream);
                        TargetEvent.Takes.Add(TargetTake);
                        TargetTake.Offset = MarginLeft;

                        // копирование некоторых атрибутов из исходного медиа
                        TargetMedia.TapeName = CurrentEvent.ActiveTake.Media.TapeName;
                        TargetMedia.Comment = CurrentEvent.ActiveTake.Media.Comment;

                        EventFileNumber++;

                    } // обход фрагментов текущего трека

                    // включить фрагменты на текущем треке
                    EventsMuteOff (CurrentTrack);

                }// обход выделенных треков


                /* ----------- + Events ---------------  */

            }
            else
            {
                filename += renderItem.Extension;
                Timecode renderStart, renderLength;
                if (renderMode == RenderMode.Selection)
                {
                    renderStart = myVegas.SelectionStart;
                    renderLength = myVegas.SelectionLength;
                }
                else
                {
                    renderStart = new Timecode();
                    renderLength = myVegas.Project.Length;
                }

                status = DoRender(filename, renderItem, renderStart, renderLength);
            }

            if (RenderStatus.Canceled == status)
                break;
        }
    }

    // perform the render.  The Render method returns a member of the
    // RenderStatus enumeration.  If it is anything other than OK,
    // exit the loops.  This will throw an error message string if the
    // render does not complete successfully.
    RenderStatus DoRender(String filePath, RenderItem renderItem, Timecode start, Timecode length)
    {
        ValidateFilePath(filePath);

        // make sure the file does not already exist
        if (!OverwriteExistingFiles && File.Exists(filePath))
        {
            throw new ApplicationException("File already exists: " + filePath);
        }

        // perform the render.  The Render method returns
        // a member of the RenderStatus enumeration.  If
        // it is anything other than OK, exit the loops.
        RenderStatus status = myVegas.Render(filePath, renderItem.Template, start, length);

        switch (status)
        {
        case RenderStatus.Complete:
        case RenderStatus.Canceled:
            break;
        case RenderStatus.Failed:
        default:
            StringBuilder msg = new StringBuilder("Render failed:\n");
            msg.Append("\n    file name: ");
            msg.Append(filePath);
            msg.Append("\n    Renderer: ");
            msg.Append(renderItem.Renderer.FileTypeName);
            msg.Append("\n    Template: ");
            msg.Append(renderItem.Template.Name);
            msg.Append("\n    Start Time: ");
            msg.Append(start.ToString());
            msg.Append("\n    Length: ");
            msg.Append(length.ToString());
            throw new ApplicationException(msg.ToString());
        }
        return status;
    }

    string FixFileName(string name)
    {
        const Char replacementChar = '-';
        foreach (char badChar in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(badChar, replacementChar);
        }
        return name;
    }

    void ValidateFilePath(string filePath)
    {
        if (filePath.Length > 260)
            throw new ApplicationException("File name too long: " + filePath);
        foreach (char badChar in Path.GetInvalidPathChars())
        {
            if (0 <= filePath.IndexOf(badChar))
            {
                throw new ApplicationException("Invalid file name: " + filePath);
            }
        }
    }

    class RenderItem
    {
        public readonly Renderer Renderer = null;
        public readonly RenderTemplate Template = null;
        public readonly string Extension = null;

        public RenderItem(Renderer r, RenderTemplate t, string e)
        {
            this.Renderer = r;
            this.Template = t;
            // need to strip off the extension's leading "*"
            if (null != e) this.Extension = e.TrimStart('*');
        }
    }

    Button BrowseButton;
    TextBox FileNameBox;
    TextBox MarginFramesBox;
    TextBox StartFileNumberBox;
    TreeView TemplateTree;
    RadioButton RenderProjectButton;
    RadioButton RenderRegionsButton;
    RadioButton RenderSelectionButton;

    Button BrowseFolderButton;
    RadioButton RenderEventsButton;
    RadioButton OfflineEventsButton;

    DialogResult ShowBatchRenderDialog()
    {
        Form dlog = new Form();
        dlog.Text = ScriptCaption;
        dlog.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
        dlog.MaximizeBox = false;
        dlog.StartPosition = FormStartPosition.CenterScreen;
        dlog.Width = 600;
        dlog.FormClosing += this.HandleFormClosing;

        int titleBarHeight = dlog.Height - dlog.ClientSize.Height;
        int buttonWidth = 80;

        //FileNameBox = AddTextControl(dlog, "Base File Name", titleBarHeight + 6, 430, 10, defaultBasePath);
        FileNameBox = AddTextControl(dlog, "", titleBarHeight + 6, 370, 10, defaultBasePath);

        BrowseButton = new Button();
        BrowseButton.Left = FileNameBox.Right + 10;
        BrowseButton.Top = FileNameBox.Top - 2;
        BrowseButton.Width = buttonWidth;
        BrowseButton.Height = BrowseButton.Font.Height + 12;
        BrowseButton.Text = "Browse...";
        BrowseButton.Click += new EventHandler(this.HandleBrowseClick);
        dlog.Controls.Add(BrowseButton);

        BrowseFolderButton = new Button();
        BrowseFolderButton.Left = BrowseButton.Right + 8;
        BrowseFolderButton.Top = FileNameBox.Top - 2;
        BrowseFolderButton.Width = buttonWidth;
        BrowseFolderButton.Height = BrowseButton.Font.Height + 12;
        BrowseFolderButton.Text = "Folder...";
        BrowseFolderButton.Click += new EventHandler(this.HandleBrowseFolderClick);
        dlog.Controls.Add(BrowseFolderButton);

        TemplateTree = new TreeView();
        TemplateTree.Left = 10;
        TemplateTree.Width = dlog.Width - 20;
        TemplateTree.Top = BrowseButton.Bottom + 10;
        TemplateTree.Height = 200;
        TemplateTree.CheckBoxes = true;
        TemplateTree.AfterCheck += new TreeViewEventHandler(this.HandleTreeViewCheck);
        dlog.Controls.Add(TemplateTree);

        int buttonTop = TemplateTree.Bottom + 25;

        int buttonsLeft = dlog.Width - (2 * (buttonWidth + 10));

        RenderProjectButton = AddRadioControl(dlog,
                                              "Render Project",
                                              25,
                                              buttonTop,
                                              true);

        RenderSelectionButton = AddRadioControl(dlog,
                                                "Render Selection",
                                                25,
                                                buttonTop + 20,
                                                (0 != myVegas.SelectionLength.Nanos));
        RenderRegionsButton = AddRadioControl(dlog,
                                              "Render Regions",
                                              25,
                                              buttonTop + 40,
                                              (0 != myVegas.Project.Regions.Count));

        RenderEventsButton = AddRadioControl(dlog,
                                             "Render Events",
                                             160,
                                             buttonTop + 10,
                                             true);
        OfflineEventsButton = AddRadioControl(dlog,
                                              "Offline Events",
                                              160,
                                              buttonTop + 30,
                                              true);

        // поле ввода количества кадров захлестов
        MarginFramesBox = AddTextControl(dlog, "Margin Frames", RenderEventsButton.Right+100, 20,
            buttonTop + 10, myDefMarginFrames.ToString());

        // поле ввода количества кадров захлестов
        StartFileNumberBox = AddTextControl(dlog, "Start File Number", RenderEventsButton.Right+100, 40,
            buttonTop + 35, myDefStartFileNumber.ToString(myDefFormatFileNumber));


        // Группа элементов Batch Render
        GroupBox BatchRenderGroupBox = new GroupBox();
        BatchRenderGroupBox.Location = new System.Drawing.Point(RenderProjectButton.Right-30, buttonTop-13);
        BatchRenderGroupBox.Size = new System.Drawing.Size(125, 80);
        BatchRenderGroupBox.Text = "Batch Render";
        BatchRenderGroupBox.SuspendLayout();
        dlog.Controls.Add(BatchRenderGroupBox);

        // Группа элементов Events
        GroupBox EventsGroupBox = new GroupBox();
        EventsGroupBox.Location = new System.Drawing.Point(OfflineEventsButton.Right-30, buttonTop-13);
        EventsGroupBox.Size = new System.Drawing.Size(280, 80);
        EventsGroupBox.Text = "Events";
        EventsGroupBox.SuspendLayout();
        dlog.Controls.Add(EventsGroupBox);

        //RenderProjectButton.Checked = true;
        RenderEventsButton.Checked = true;

        Button okButton = new Button();
        okButton.Text = "OK";
        okButton.Left = dlog.Width - (buttonWidth+20);
        okButton.Top = buttonTop - 10;
        okButton.Width = buttonWidth;
        okButton.Height = okButton.Font.Height + 12;
        okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
        dlog.AcceptButton = okButton;
        dlog.Controls.Add(okButton);

        Button cancelButton = new Button();
        cancelButton.Text = "Cancel";
        cancelButton.Left = okButton.Left;
        cancelButton.Top =  okButton.Top + 30;
        cancelButton.Height = cancelButton.Font.Height + 12;
        cancelButton.Width = buttonWidth;
        cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
        dlog.CancelButton = cancelButton;
        dlog.Controls.Add(cancelButton);

        Button AboutButton = new Button();
        AboutButton.Text = "About";
        AboutButton.Left = okButton.Left;
        AboutButton.Top =  cancelButton.Top + 30;
        AboutButton.Width = buttonWidth;
        AboutButton.Height = okButton.Font.Height + 12;
        AboutButton.Click += new EventHandler(this.HandleAboutClick);
        dlog.Controls.Add(AboutButton);

        dlog.Height = titleBarHeight + AboutButton.Bottom + 10;
        dlog.ShowInTaskbar = false;

        FillTemplateTree();

        return dlog.ShowDialog(myVegas.MainWindow);
    }

    TextBox AddTextControl(Form dlog, string labelName, int left, int width, int top, string defaultValue)
    {

        TextBox textbox = new TextBox();
        textbox.Multiline = false;
        textbox.Left = left;
        textbox.Top = top;
        textbox.Width = width;
        textbox.Text = defaultValue;
        dlog.Controls.Add(textbox);

        Label label = new Label();
        label.AutoSize = true;
        label.Text = " " + labelName;
        label.Left = textbox.Right;
        label.Top = top + 4;
        dlog.Controls.Add(label);

        return textbox;
    }

    RadioButton AddRadioControl(Form dlog, string labelName, int left, int top, bool enabled)
    {


        RadioButton radiobutton = new RadioButton();
        radiobutton.Left = left;
        radiobutton.Width = 20;
        radiobutton.Top = top;
        radiobutton.Enabled = enabled;
        dlog.Controls.Add(radiobutton);

        Label label = new Label();
        label.AutoSize = true;
        label.Text = labelName;
        label.Left = radiobutton.Right;
        label.Top = top + 4;
        label.Enabled = enabled;
        dlog.Controls.Add(label);

        return radiobutton;
    }

    void FillTemplateTree()
    {
        int projectAudioChannelCount = 0;
        if (AudioBusMode.Stereo == myVegas.Project.Audio.MasterBusMode)
        {
            projectAudioChannelCount = 2;
        }
        else if (AudioBusMode.Surround == myVegas.Project.Audio.MasterBusMode)
        {
            projectAudioChannelCount = 6;
        }
        bool projectHasVideo = ProjectHasVideo();
        bool projectHasAudio = ProjectHasAudio();
        foreach (Renderer renderer in myVegas.Renderers)
        {
            try
            {
                string rendererName = renderer.FileTypeName;
                TreeNode rendererNode = new TreeNode(rendererName);
                rendererNode.Tag = new RenderItem(renderer, null, null);
                foreach (RenderTemplate template in renderer.Templates)
                {
                    try
                    {
                        // filter out invalid templates
                        if (!template.IsValid())
                        {
                            continue;
                        }
                        // filter out video templates when project has
                        // no video.
                        if (!projectHasVideo && (0 < template.VideoStreamCount))
                        {
                            continue;
                        }
                        // filter out audio-only templates when project has no audio
                        if (!projectHasAudio && (0 == template.VideoStreamCount) && (0 < template.AudioStreamCount))
                        {
                            continue;
                        }
                        // filter out templates that have more channels than the project
                        if (projectAudioChannelCount < template.AudioChannelCount)
                        {
                            continue;
                        }
                        // filter out templates that don't have
                        // exactly one file extension
                        string[] extensions = template.FileExtensions;
                        if (1 != extensions.Length)
                        {
                            continue;
                        }
                        string templateName = template.Name;
                        TreeNode templateNode = new TreeNode(templateName);

                        if (String.Equals(templateName, myDefRenderTempl))
                        {
                            templateNode.ToolTipText = "Default Template";
                            templateNode.BackColor = Color.FromArgb(255, 218, 179, 179);
                            templateNode.Checked = true;
                            rendererNode.Checked = true;
                            rendererNode.Expand();
                        }

                        templateNode.Tag = new RenderItem(renderer, template, extensions[0]);
                        rendererNode.Nodes.Add(templateNode);
                    }
                    catch (Exception e)
                    {
                        // skip it
                        MessageBox.Show(e.ToString());
                    }
                }
                if (0 == rendererNode.Nodes.Count)
                {
                    continue;
                }
                else if (1 == rendererNode.Nodes.Count)
                {
                    // skip it if the only template is the project
                    // settings template.
                    if (0 == ((RenderItem) rendererNode.Nodes[0].Tag).Template.Index)
                    {
                        continue;
                    }
                }
                else
                {
                    TemplateTree.Nodes.Add(rendererNode);
                }
            }
            catch
            {
                // skip it
            }
        }
    }

    bool ProjectHasVideo()
    {
        foreach (Track track in myVegas.Project.Tracks)
        {
            if (track.IsVideo())
            {
                return true;
            }
        }
        return false;
    }

    bool ProjectHasAudio()
    {
        foreach (Track track in myVegas.Project.Tracks)
        {
            if (track.IsAudio())
            {
                return true;
            }
        }
        return false;
    }

    void UpdateSelectedTemplates()
    {
        SelectedTemplates.Clear();
        foreach (TreeNode node in TemplateTree.Nodes)
        {
            foreach (TreeNode templateNode in node.Nodes)
            {
                if (templateNode.Checked)
                {
                    SelectedTemplates.Add(templateNode.Tag);
                }
            }
        }
    }

    void HandleBrowseClick(Object sender, EventArgs args)
    {
        SaveFileDialog saveFileDialog = new SaveFileDialog();
        saveFileDialog.Filter = "All Files (*.*)|*.*";
        saveFileDialog.CheckPathExists = true;
        saveFileDialog.AddExtension = false;
        if (null != FileNameBox)
        {
            string filename = FileNameBox.Text;
            string initialDir = Path.GetDirectoryName(filename);
            if (Directory.Exists(initialDir))
            {
                saveFileDialog.InitialDirectory = initialDir;
            }
            saveFileDialog.DefaultExt = Path.GetExtension(filename);
            saveFileDialog.FileName = Path.GetFileNameWithoutExtension(filename);
        }
        if (System.Windows.Forms.DialogResult.OK == saveFileDialog.ShowDialog())
        {
            if (null != FileNameBox)
            {
                FileNameBox.Text = Path.GetFullPath(saveFileDialog.FileName);
            }
        }
    }


    void HandleBrowseFolderClick(Object sender, EventArgs args)
    {

        FolderBrowserDialog BrowseFolderDialog = new FolderBrowserDialog();

        if (null != FileNameBox)
        {
            string filename = FileNameBox.Text;
            string initialDir = Path.GetDirectoryName(filename);
            if (Directory.Exists(initialDir))
                BrowseFolderDialog.SelectedPath = initialDir;
        }

        if (System.Windows.Forms.DialogResult.OK == BrowseFolderDialog.ShowDialog())
        {
            if (null != FileNameBox)
            {
                FileNameBox.Text = Path.GetFullPath(BrowseFolderDialog.SelectedPath) + Path.DirectorySeparatorChar;
            }
        }

    }


    void HandleTreeViewCheck(object sender, TreeViewEventArgs args)
    {
        if (args.Node.Checked)
        {
            if (0 != args.Node.Nodes.Count)
            {
                if ((args.Action == TreeViewAction.ByKeyboard) || (args.Action == TreeViewAction.ByMouse))
                {
                    SetChildrenChecked(args.Node, true);
                }
            }
            else if (!args.Node.Parent.Checked)
            {
                args.Node.Parent.Checked = true;
            }
        }
        else
        {
            if (0 != args.Node.Nodes.Count)
            {
                if ((args.Action == TreeViewAction.ByKeyboard) || (args.Action == TreeViewAction.ByMouse))
                {
                    SetChildrenChecked(args.Node, false);
                }
            }
            else if (args.Node.Parent.Checked)
            {
                if (!AnyChildrenChecked(args.Node.Parent))
                {
                    args.Node.Parent.Checked = false;
                }
            }
        }
    }

    void HandleFormClosing(Object sender, FormClosingEventArgs args)
    {
        Form dlg = sender as Form;
        if (null == dlg) return;
        if (DialogResult.OK != dlg.DialogResult) return;
        string outputFilePath = FileNameBox.Text;
        try
        {
            string outputDirectory = Path.GetDirectoryName(outputFilePath);
            if (!Directory.Exists(outputDirectory)) throw new ApplicationException();
        }
        catch
        {
            string title = "Invalid Directory";
            StringBuilder msg = new StringBuilder();
            msg.Append("The output directory does not exist.\n");
            msg.Append("Please specify the directory and base file name using the Browse button.");
            MessageBox.Show(dlg, msg.ToString(), title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            args.Cancel = true;
            return;
        }
        try
        {
            string baseFileName = Path.GetFileName(outputFilePath);
            if (String.IsNullOrEmpty(baseFileName) && (false == RenderEventsButton.Checked && false == OfflineEventsButton.Checked) )
                throw new ApplicationException();
            if (-1 != baseFileName.IndexOfAny(Path.GetInvalidFileNameChars()))
                throw new ApplicationException();
        }
        catch
        {
            string title = "Invalid Base File Name";
            StringBuilder msg = new StringBuilder();
            msg.Append("The base file name is not a valid file name.\n");
            msg.Append("Make sure it contains one or more valid file name characters.");
            MessageBox.Show(dlg, msg.ToString(), title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            args.Cancel = true;
            return;
        }
        UpdateSelectedTemplates();
        if (0 == SelectedTemplates.Count)
        {
            string title = "No Templates Selected";
            StringBuilder msg = new StringBuilder();
            msg.Append("No render templates selected.\n");
            msg.Append("Select one or more render templates from the available formats.");
            MessageBox.Show(dlg, msg.ToString(), title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            args.Cancel = true;
            return;
        }
    }

    void SetChildrenChecked(TreeNode node, bool checkIt)
    {
        foreach (TreeNode childNode in node.Nodes)
        {
            if (childNode.Checked != checkIt)
                childNode.Checked = checkIt;
        }
    }

    bool AnyChildrenChecked(TreeNode node)
    {
        foreach (TreeNode childNode in node.Nodes)
        {
            if (childNode.Checked) return true;
        }
        return false;
    }


    //
    //
    //
    int StringToInt(string InputString, int DefaultInt)
    {
        int OutInt;
        Regex regex = new Regex("[^0-9]");
        InputString = regex.Replace(InputString, "");

        if (false == Int32.TryParse(InputString, out OutInt))
        {
            OutInt = DefaultInt;
        }

        return OutInt;
    }


    //
    //
    //
    ArrayList GetSelectedTrack(MediaType mediaType)
    {
        ArrayList SelectedTracks = new ArrayList();

        foreach (Track track in myVegas.Project.Tracks)
        {
            if (true == track.Selected && mediaType == track.MediaType)
            {
                SelectedTracks.Add(track);
            }
        }
        return SelectedTracks;
    }


    //
    //
    //
    int GetEventsCount (ArrayList Tracks)
    {
        int EventsCount = 0;

        // обход треков
        foreach (Track CurrentTrack in Tracks)
            EventsCount += CurrentTrack.Events.Count;

        return EventsCount;
    }



    //
    //
    //
    Dictionary<string, string> PresetsFromFile (Dictionary<string, string> ExtRenderCommand, string FilePath)
    {

        String scriptDirectory = Path.GetDirectoryName(Sony.Vegas.Script.File);
        String FileFullPath = Path.Combine(scriptDirectory, FilePath);

        // если файл существует
        if (File.Exists(FileFullPath))
        {
            // содержимое файла
            string text = System.IO.File.ReadAllText(FileFullPath);

            // удаление лишнего
            ArrayList RegArrayList = new ArrayList();
            RegArrayList.Add(@"^\s+|\s+$");
            RegArrayList.Add(@"(//|#).*");
            //RegArrayList.Add(@"(\r?\n){2,}");

            foreach (string RegexItem in RegArrayList)
            {
                Regex regex = new Regex(RegexItem, RegexOptions.Multiline);
                text = regex.Replace(text, "");
            }


            // разбивка содержимого файла на строки
            char[] determine = new char[] { '\r', '\n' };
            string[] lines = text.Split(determine, StringSplitOptions.RemoveEmptyEntries);

            // привести к четному количеству строк
            int length = (lines.Length / 2) * 2;

            // добавление строк в словарь
            for (int i = 0; i < length; i++)
            {
               // четная строка
                if (i % 2 == 0)
                {
                    // если ключ еще не существует
                    if (false == ExtRenderCommand.ContainsKey(lines[i]))
                    {
                        ExtRenderCommand.Add(lines[i], lines[i+1]);
                    }
               }
            }
        }

        return ExtRenderCommand;
    }


    //
    //
    //
    SaveFileDialog SaveBatchFileDialog (Dictionary<string, string> ExtRenderCommand, string BatFileName)
    {

        // Список ключей словаря
        List<string> KeysList = new List<string>(ExtRenderCommand.Keys);

        SaveFileDialog saveFileDialog = new SaveFileDialog();
        saveFileDialog.Filter = String.Join("|", KeysList.ToArray());
        saveFileDialog.FilterIndex  = 4;
        saveFileDialog.CheckPathExists = true;
        saveFileDialog.AddExtension = true;
        saveFileDialog.DefaultExt = Path.GetExtension(BatFileName);
        saveFileDialog.FileName = Path.GetFileNameWithoutExtension(BatFileName);
        saveFileDialog.Title = "Save Batch File";

        if (null != FileNameBox)
        {
            string filename = FileNameBox.Text;
            string initialDir = Path.GetDirectoryName(filename);
            if (Directory.Exists(initialDir))
            {
                saveFileDialog.InitialDirectory = initialDir;
            }
        }

        return saveFileDialog;
    }


    //
    //
    //
    string BatFileContent(string Placeholder, ArrayList ListOptions)
    {
        //
        StringBuilder BatFileContent = new StringBuilder("");
        BatFileContent.AppendLine();

        foreach (string LineOptions in ListOptions)
        {
            string[] Options = LineOptions.Split(new Char[] {'|'});

            //
            Options[2] = Options[2].Replace(",", ".");
            Options[3] = Options[3].Replace(",", ".");

            // строка команды
            string CommandLine = String.Format(Placeholder, Options[0], Options[1], Options[2], Options[3]);
            // добавление в bat-файл
            BatFileContent.Append(CommandLine);
            BatFileContent.AppendLine();
        }
        return BatFileContent.ToString();
    }


    //
    //
    //
    void CopyFades (TrackEvent SourceEvent, TrackEvent TargetEvent)
    {

        // копирование затуханий
        TargetEvent.FadeIn.Length = SourceEvent.FadeIn.Length;
        TargetEvent.FadeIn.Curve = SourceEvent.FadeIn.Curve;
        TargetEvent.FadeIn.Gain = SourceEvent.FadeIn.Gain;
        TargetEvent.FadeIn.ReciprocalCurve = SourceEvent.FadeIn.ReciprocalCurve;

        TargetEvent.FadeOut.Length = SourceEvent.FadeOut.Length;
        TargetEvent.FadeOut.Curve = SourceEvent.FadeOut.Curve;
        TargetEvent.FadeOut.Gain = SourceEvent.FadeOut.Gain;
        TargetEvent.FadeOut.ReciprocalCurve = SourceEvent.FadeOut.ReciprocalCurve;

        // копирование переходов (пресеты не поддерживаются)
        String PluginName;
        PlugInNode Plugin;

        Effect TransitionIn = (Effect) SourceEvent.FadeIn.Transition;
        if (null != TransitionIn)
        {
            PluginName = TransitionIn.Description;
            Plugin = myVegas.Transitions.GetChildByName(PluginName);
            TargetEvent.FadeIn.Transition = new Effect(Plugin);
        }

        Effect TransitionOut = (Effect) SourceEvent.FadeOut.Transition;
        if (null != TransitionOut)
        {
            PluginName = TransitionOut.Description;
            Plugin = myVegas.Transitions.GetChildByName(PluginName);
            TargetEvent.FadeIn.Transition = new Effect(Plugin);
        }

    }


    //
    //
    //
    void EventsMuteOn (Track CurrentTrack, TrackEvent CurrentEvent)
    {

        // если первый фрагметн трека
        if (0 == CurrentEvent.Index)
        {
            // обход фрагментов трека
            foreach (TrackEvent Event in CurrentTrack.Events)
                Event.Mute = true;

            // если последний фрагмент трека
        }
        else if (CurrentEvent.Index == CurrentTrack.Events.Count - 1)
        {
            CurrentTrack.Events[CurrentEvent.Index - 1].Mute = true;
        }
        else
        {
            CurrentTrack.Events[CurrentEvent.Index - 1].Mute = true;
            CurrentTrack.Events[CurrentEvent.Index + 1].Mute = true;
        }

        CurrentEvent.Mute = false;
    }


    //
    //
    //
    void EventsMuteOff (Track CurrentTrack)
    {

        // обход фрагментов трека
        foreach (TrackEvent Event in CurrentTrack.Events)
            Event.Mute = false;

    }


    //
    //
    //
    void HandleAboutClick(Object sender, EventArgs args)
    {
        string Message = ScriptCaption + "\n" +
            "Version: " + ScriptVersion + "\n\n" +
            "Modifed by Chainick Dmitry <chainick@narod.ru>";


        MessageBox.Show(
            Message,
            ScriptCaption,
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }
}
