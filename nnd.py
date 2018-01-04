import pandas as pd
import numpy as np
import xlrd
import glob
import csv

filename="Raw data-105 shoal data-Trial     1.xlsx"

class Fish(object):
	def __init__(self, df):
		self.df = pd.DataFrame(data=df, copy=True)
		self.trialid = self.df.loc[df["Number of header lines:"] == "Trial name"]["35"].item().split()[1]
		self.arenaid = self.df.loc[self.df["Number of header lines:"] == "Arena name"]["35"].item().split()[1]
		self.shoalid = self.trialid + self.arenaid
		self.group = self.df.loc[self.df["Number of header lines:"] == "trt"]["35"].item()
		self.subjectid = self.df.loc[self.df["Number of header lines:"] == "Subject name"]["35"].item()
		self.headeridx = self.df[self.df["Number of header lines:"] == "Trial time"].index.item()
		self.parse(self.df)
	def parse(self, messy):
		messy.columns = messy.iloc[self.headeridx]
		messy.drop(messy.index[0:(self.headeridx+2)], inplace=True)
		exclude, keep = [], []
		for i in messy.columns.tolist():
			if "Recording time" in i:
				keep.append("Time")
			elif "Subject" in i and self.subjectid in i:
				exclude.append(i)
			elif "Subject" in i:
				keep.append(self.subjectid.split()[1] + "to" + i.split()[5])
			else:
				exclude.append(i)
		for i in exclude:
			messy.drop(i, axis=1, inplace=True)
		messy.columns = keep
		messy.reset_index(drop=True, inplace=True)
		messy.replace("-", np.NaN, inplace=True)
		messy["NND"] = messy[messy.columns[1:]].min(axis=1)
		messy.set_index(["Time"], drop=True, inplace=True)
		messy.drop(messy.columns[:3], axis=1, inplace=True)

def compile(shoal):


imp = pd.ExcelFile(filename)
sheets = imp.sheet_names

fish = []
shoalcount = []

for sheet in sheets:
	fish.append(Fish(imp.parse(sheet)))