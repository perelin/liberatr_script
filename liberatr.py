import win32com.client
#import eyed3
import logging
from pprint import pprint
import json
from datetime import datetime
#import timeit
#from mutagen.mp3 import MP3
from mutagen.id3 import ID3, POPM
#from mutagen.easyid3 import EasyID3
import argparse


# init 


itunes = win32com.client.Dispatch("iTunes.Application")
mainLibrary = itunes.LibraryPlaylist
libraryTracks = mainLibrary.Tracks
logging.basicConfig(filename="itunesoperations.log.txt",level=logging.WARNING)
now = datetime.now()
logging.critical('###### Starting new Script run @ %s' % (str(now)))
popmUser = "LiberatR"

def walkThroughItunesLib():
	
	for track in libraryTracks:
		if track.Rating > 0:
			#audiofile = eyed3.load(track.Location)
			#print(audiofile.tag.version,itunes.ITObjectPersistentIDHigh(track), itunes.ITObjectPersistentIDLow(track), track.Location)
			trackID3 = ID3(track.Location)
			popFrames = trackID3.getall('POPM')
			print(track.Artist, track.Name)
			print(track.Rating, popFrames)
			#print(popFrames[0].email)
			
def testLibs():
	
	#track = libraryTracks.ItemByPersistentID(1317396620,-33136939) # Perfect Circle - Judith v 2.3.0
	#track = libraryTracks.ItemByPersistentID(660927261,1657575963) # Celentano v 2.4.0
	track = libraryTracks.ItemByPersistentID(1371381314,1577795262) # Coulton v 2.4.0
	#print(track.Artist, track.Name, track.Album, track.Location)
	
	# iTunes
	#print(track.Comment)
	#track.Comment = 'test'
	#sprint(track.Comment)
	
	# eyeD3
	#trackFile = eyed3.load(track.Location)
	#print(trackFile.tag.artist)
	#pprint(type(str(trackFile.tag.version)))
	#tag = eyed3.Tag()
	#tag.link(track.Location)
	
	# Mutagen
	
	#trackFile = MP3(track.Location)
	#trackID3 = ID3(track.Location)
	trackID3 = ID3(u"E:\\00_audio\\Adriano Celentano\\Super Best\\20 Susanna.mp3")
	
	#print(trackFile.info.length, trackFile.info.bitrate)
	#print(dir(trackFile))
	#trackID3.pprint()
	#pprint(trackID3)
	#trackFile.pprint()
	#pprint(trackFile)
	#print EasyID3.valid_keys.keys()
	#print(trackID3.version, __name__)
	#pprint(trackID3)
	#trackID3.add(POPM(email="test@foo.com", rating=222, count=12))
	#trackID3.save()
	#pprint(trackID3)
	#trackID3.delall('POPM')
	print(trackID3.getall('POPM'))

def saveAllItunesRatingsToPOPM():

	for track in libraryTracks:
		if track.Rating > 0:
			trackID3 = ID3(track.Location)
			popmRating = convertItunesRatingToPOPM(track.Rating)
			trackID3.add(POPM(email=popmUser, rating=popmRating, count=track.PlayedCount))
			trackID3.save()

def restoreAllItunesRatingsfromPOPM():

	for track in libraryTracks:
		popmRating = getPOPMRatingFromFile(track.Location)
		if popmRating:
			itunesRating = convertPOPMRatingToItunes(popmRating)
			if itunesRating > track.Rating:
				track.Rating = itunesRating
		
def getPOPMRatingFromFile(fileLocation):
		trackID3 = ID3(fileLocation)
		trackPOPMFrames = trackID3.getall("POPM")
		for frame in trackPOPMFrames:
			if frame.email == popmUser:
				return frame.rating

def convertPOPMRatingToItunes(popmRating):

	itunesRating = 0
	
	if popmRating > 0 and popmRating <= 1:
		itunesRating = 20
	elif popmRating <= 64:
		itunesRating = 40
	elif popmRating <= 128:
		itunesRating = 60
	elif popmRating <= 196:
		itunesRating = 80
	elif popmRating <= 255:
		itunesRating = 100
	
	return itunesRating
	
def convertItunesRatingToPOPM(itunesRating):
	
	popmRating = 0
	
	if itunesRating > 0 and itunesRating <= 20:
		popmRating = 1
	elif itunesRating <= 40:
		popmRating = 64
	elif itunesRating <= 60:
		popmRating = 128
	elif itunesRating <= 80:
		popmRating = 196
	elif itunesRating <= 100:
		popmRating = 255
	
	return popmRating

def removePOPMFramesFromLibrary():
	for track in libraryTracks:
		removePOPMFramesFromFile(track.Location)
	
def removePOPMFramesFromFile(fileLocation):
	trackID3 = ID3(fileLocation)
	if trackID3.getall('POPM'):
		trackID3.delall('POPM')
		trackID3.save()
	
def getTagVersionsInLib():
	
	tagVersions = { (1,1,0):0, (1,0,0):0, (2,3,0):0, (2,4,0):0, (2,2,0):0, (0,0,0):0}
	
	for track in libraryTracks:
		if track.Rating > 0:
			#print(itunes.ITObjectPersistentIDHigh(track), itunes.ITObjectPersistentIDLow(track), track.Location)
	
			try:
				#audiofile = eyed3.load(track.Location)
				trackID3 = ID3(track.Location)
				tagVersions[trackID3.version] += 1
				#print(audiofile.tag.version,track.Location)
			except:
				tagVersions[(0,0,0)] += 1
			
	pprint(tagVersions)
	
def cleanUnicode(unicodeData):
	return unicodeData.encode('ascii', 'ignore')
	
# Run
parser = argparse.ArgumentParser(description="Works with iTunes and POPM ratings")
parser.add_argument("mode", help="""
					rm = remove all ratings, 
					ls = show rated songs in iTunes lib, 
					cp = copy itunes ratings to POPM ID3 frame,
					vs = get an overview of all ID3 tag version in iTunes lib,
					mv = move POPM ratings to itunes
					""")
args = parser.parse_args()

if args.mode == "ls":
	walkThroughItunesLib()
elif args.mode == "rm":
	removePOPMFramesFromLibrary()
elif args.mode == "cp":
	saveAllItunesRatingsToPOPM()
elif args.mode == "vs":
	getTagVersionsInLib()
elif args.mode == "mv":
	restoreAllItunesRatingsfromPOPM()
	