import System;
import System.IO;
import System.Windows.Forms;
import Microsoft.Win32;
import Sony.Vegas;
var evnt: TrackEvent;
var myRegion: Region;
var RegionNumber;
var RegionName: String;


try {
    //Find the selected event
    var track = FindSelectedTrack();
    if (null == track)
        throw "no selected track";

    var eventEnum = new Enumerator(track.Events);
    while (!eventEnum.atEnd()) {
        evnt = TrackEvent(eventEnum.item());
        var MyFilePath = evnt.ActiveTake.MediaPath;
        var extFileName = Path.GetFileName(MyFilePath);
        var baseFileName = Path.GetFileNameWithoutExtension(extFileName); // Media file name for this event
        var nameR = "null";

        myRegion = new Region(evnt.Start, evnt.Length, track.Name); //Insert a region over this event
        Vegas.Project.Regions.Add(myRegion);
        eventEnum.moveNext();
        RegionNumber++;
    }
} catch (e) {
    MessageBox.Show(e);
}

function FindSelectedTrack(): Track {
    var trackEnum = new Enumerator(Vegas.Project.Tracks);
    while (!trackEnum.atEnd()) {
        var track: Track = Track(trackEnum.item());
        if (track.Selected) {
            return track;
        }
        trackEnum.moveNext();
    }
    return null;
}