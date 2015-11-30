#!/usr/bin/env python
#-*- coding: utf-8 -*-

#phenol - 
#This application is free software; you can redistribute
#it and/or modify it under the terms of the GNU General Public License
#defined in the COPYING file

#2010 Charlie Barnes.

import sys
import os
import gtk
import gobject
import mimetypes
import xlrd
from pygtk_chart import bar_chart

class phenolActions():
    def __init__(self):

        #Load the widget tree
        builder = ""
        self.builder = gtk.Builder()
        self.builder.add_from_string(builder, len(builder))
        self.builder.add_from_file("ui.xml")

        signals = {
                   "mainQuit":self.main_quit,
                   "showLicense":self.show_about_dialog,
                   "selectFile":self.select_file,
                   "calculate":self.calculate,
                   "parse":self.parse,
                  }
        self.builder.connect_signals(signals)
        self.builder.get_object("window1").show()
        self.taxa = { }

    def parse(self, filename):
 
        if self.builder.get_object("viewport1").get_child():
            self.builder.get_object("viewport1").remove(self.builder.get_object("viewport1").get_child())
                
        if self.builder.get_object("eventbox1").get_child():
            self.builder.get_object("eventbox1").get_child().destroy()
 
        self.taxa = { }
 
        cursor = gtk.gdk.Cursor(gtk.gdk.WATCH)
        self.builder.get_object("window1").window.set_cursor(cursor)
    
        while gtk.events_pending():
            gtk.main_iteration()
                    
        filetype = mimetypes.guess_type(filename)[0]
        
        if filetype == "application/vnd.ms-excel":
            book = xlrd.open_workbook(filename)
            
            if book.nsheets > 1:
                
                dialog = self.builder.get_object("dialog1")

                try:
                    self.builder.get_object("hbox5").get_children()[1].destroy()           
                except IndexError:
                    pass
                    
                combobox = gtk.combo_box_new_text()
                
                for name in book.sheet_names():
                    combobox.append_text(name)
                    
                combobox.set_active(0)
                combobox.show()
                self.builder.get_object("hbox5").add(combobox)
                
                self.builder.get_object("window1").window.set_cursor(None)
            
                while gtk.events_pending():
                    gtk.main_iteration()
                
                response = dialog.run()

                if response == 1:
                    sheet = book.sheet_by_name(combobox.get_active_text())
                else:
                    dialog.hide()
                    return -1
                    
                dialog.hide()
                
            else:
                sheet = book.sheet_by_index(0)

            cursor = gtk.gdk.Cursor(gtk.gdk.WATCH)
            self.builder.get_object("window1").window.set_cursor(cursor)
        
            while gtk.events_pending():
                gtk.main_iteration()
        
            for col_index in range(sheet.ncols):
                if sheet.cell(0, col_index).value == "Species":
                    taxon_position = col_index
                elif sheet.cell(0, col_index).value == "Taxon":
                    taxon_position = col_index
                elif sheet.cell(0, col_index).value == "Taxon Name":
                    taxon_position = col_index
                elif sheet.cell(0, col_index).value == "Date":
                    date_position = col_index
            
            combobox_taxa = []  
                  
            self.taxa["all records"] = {"Jan": 0,
                                        "Feb": 0,
                                        "Mar": 0,
                                        "Apr": 0,
                                        "May": 0,
                                        "Jun": 0,
                                        "Jul": 0,
                                        "Aug": 0,
                                        "Sep": 0, 
                                        "Oct": 0,
                                        "Nov": 0,
                                        "Dec": 0,
                                       }
                                       
            for row_index in range(1, sheet.nrows):
                taxon = sheet.cell(row_index, taxon_position).value
                date = sheet.cell(row_index, date_position).value
                    
                date_split = None
                
                try:    
                    xlrddate = xlrd.xldate_as_tuple(date, book.datemode)
                    month = xlrddate[1]
                except ValueError:
                    date_split = None
                    
                    # try some standard dates  
                    if date.count('/') == 2:  
                        date_split = date.split('/')
                    elif date.count('-') == 2:  
                        date_split = date.split('-')
                    elif date.count('\\') == 2:  
                        date_split = date.split('\\')
                    
                    if date_split:                      # british only please!     
                        if int(date_split[0]) > 31:     #y/m/d
                            month = int(date_split[1])
                        elif int(date_split[2]) > 31:   #d/m/y
                            month = int(date_split[1])
                
                if month == 1:
                    month = "Jan"
                elif month == 2:
                    month = "Feb"
                elif month == 3:
                    month = "Mar"
                elif month == 4:
                    month = "Apr"
                elif month == 5:
                    month = "May"
                elif month == 6:
                    month = "Jun"
                elif month == 7:
                    month = "Jul"
                elif month == 8:
                    month = "Aug"
                elif month == 9:
                    month = "Sep"
                elif month == 10:
                    month = "Oct"
                elif month == 11:
                    month = "Nov"
                elif month == 12:
                    month = "Dec"
                    
                if not self.taxa.has_key(taxon.lower()):
                    combobox_taxa.append(taxon)
                    
                    self.taxa[taxon.lower()] = {"Jan": 0,
                                                "Feb": 0,
                                                "Mar": 0,
                                                "Apr": 0,
                                                "May": 0,
                                                "Jun": 0,
                                                "Jul": 0,
                                                "Aug": 0,
                                                "Sep": 0, 
                                                "Oct": 0,
                                                "Nov": 0,
                                                "Dec": 0,
                                               }


                self.taxa[taxon.lower()][month] = self.taxa[taxon.lower()][month] + 1
                self.taxa["all records"][month] = self.taxa["all records"][month] + 1


        cursor = gtk.gdk.Cursor(gtk.gdk.WATCH)
        self.builder.get_object("window1").window.set_cursor(cursor)
    
        while gtk.events_pending():
            gtk.main_iteration()
            
        combobox = gtk.combo_box_entry_new_text()
        combobox.connect('changed', self.calculate)        
        combobox_taxa.sort()
        combobox.append_text("All records")
        
        for item in combobox_taxa:
            combobox.append_text(item)
            
        self.builder.get_object("eventbox1").add(combobox)
        combobox.set_wrap_width(3)

        combobox.show()
            
        self.builder.get_object("window1").window.set_cursor(None)
        
        while gtk.events_pending():
            gtk.main_iteration()
                             
    def calculate(self, widget):

        combobox = self.builder.get_object("eventbox1").get_child()
        model = combobox.get_model()
        
        taxon = combobox.get_child().get_text()

        if self.builder.get_object("viewport1").get_child():
            self.builder.get_object("viewport1").get_child().destroy()
                
        if taxon.lower() in self.taxa:
            chart = bar_chart.BarChart()

            total_rec = 0

            for month in ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",]:
                bar = bar_chart.Bar(month, self.taxa[taxon.lower()][month], month)
                chart.add_bar(bar)
                total_rec = total_rec + self.taxa[taxon.lower()][month]
                
            if total_rec == 1:
                s_ = ''
            else:
                s_ = 's'
                
            chart.title.set_text(''.join(['Temporal distribution of ', taxon, ' (', str(total_rec), ' record', s_, ')']))

            self.builder.get_object("viewport1").add(chart)
            chart.show()
    
    def select_file(self, widget):
        filetype = mimetypes.guess_type(self.builder.get_object("filechooserbutton2").get_filename())[0]
        
        if filetype == "application/vnd.ms-excel":
            self.parse(self.builder.get_object("filechooserbutton2").get_filename())
                          
    def main_quit(self, widget, var=None):
        gtk.main_quit()
       
    def show_about_dialog(self, widget):
       about=gtk.AboutDialog()
       about.set_name("phenol")
       about.set_copyright("2010 Charlie Barnes")
       about.set_authors(["Charlie Barnes <charlie@cucaera.co.uk>"])
       about.set_license("phenol is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the Licence, or (at your option) any later version.\n\nphenol is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.\n\nYou should have received a copy of the GNU General Public License along with phenol; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA")
       about.set_wrap_license(True)
       about.set_website("http://cucaera.co.uk/software/phenol/")
       about.set_transient_for(self.builder.get_object("window1"))
       result=about.run()
       about.destroy()

if __name__ == '__main__':
    phenolActions()
    gtk.main()
    
