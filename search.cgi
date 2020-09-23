#!/usr/bin/perl
##############################################################################
# Simple Search                 Version 1.0                                  #
# Copyright 1996 Matt Wright    mattw@worldwidemart.com                      #
# Created 12/16/95              Last Modified 12/16/95                       #
# Scripts Archive at:           http://www.worldwidemart.com/scripts/        #
##############################################################################
# COPYRIGHT NOTICE                                                           #
# Copyright 1996 Matthew M. Wright  All Rights Reserved.                     #
#                                                                            #
# Simple Search may be used and modified free of charge by anyone so long as #
# this copyright notice and the comments above remain intact.  By using this #
# code you agree to indemnify Matthew M. Wright from any liability that      #  
# might arise from it's use.                                                 #  
#                                                                            #
# Selling the code for this program without prior written consent is         #
# expressly forbidden.  In other words, please ask first before you try and  #
# make money off of my program.                                              #
#                                                                            #
# Obtain permission before redistributing this software over the Internet or #
# in any other medium.  In all cases copyright and header must remain intact.#
##############################################################################
# Define Variables							     #

$basedir = '../../';
$baseurl = '../../';
@files = ('*.html','*.htm','*.txt');
$title = "Elucid Software Search Results";
$title_url = 'http://elucidsoftware.hypermart.net/index.htm';
$search_url = 'http://elucidsoftware.hypermart.net';

# Done									     #
##############################################################################

# Parse Form Search Information
&parse_form;

# Get Files To Search Through
&get_files;

# Search the files
&search;

# Print Results of Search
&return_html;


sub parse_form {

   # Get the input
   read(STDIN, $buffer, $ENV{'CONTENT_LENGTH'});

   # Split the name-value pairs
   @pairs = split(/&/, $buffer);

   foreach $pair (@pairs) {
      ($name, $value) = split(/=/, $pair);

      $value =~ tr/+/ /;
      $value =~ s/%([a-fA-F0-9][a-fA-F0-9])/pack("C", hex($1))/eg;

      $FORM{$name} = $value;
   }
}

sub get_files {

   chdir($basedir);
   foreach $file (@files) {
      $ls = `ls $file`;
      @ls = split(/\s+/,$ls);
      foreach $temp_file (@ls) {
         if (-d $file) {
            $filename = "$file$temp_file";
            if (-T $filename) {
               push(@FILES,$filename);
            }
         }
         elsif (-T $temp_file) {
            push(@FILES,$temp_file);
         }
      }
   }
}

sub search {

   @terms = split(/\s+/, $FORM{'terms'});

   foreach $FILE (@FILES) {

      open(FILE,"$FILE");
      @LINES = <FILE>;
      close(FILE);

      $string = join(' ',@LINES);
      $string =~ s/\n//g;
      if ($FORM{'boolean'} eq 'AND') {
         foreach $term (@terms) {
            if ($FORM{'case'} eq 'Insensitive') {
               if (!($string =~ /$term/i)) {
                  $include{$FILE} = 'no';
  		  last;
               }
               else {
                  $include{$FILE} = 'yes';
               }
            }
            elsif ($FORM{'case'} eq 'Sensitive') {
               if (!($string =~ /$term/)) {
                  $include{$FILE} = 'no';
                  last;
               }
               else {
                  $include{$FILE} = 'yes';
               }
            }
         }
      }
      elsif ($FORM{'boolean'} eq 'OR') {
         foreach $term (@terms) {
            if ($FORM{'case'} eq 'Insensitive') {
               if ($string =~ /$term/i) {
                  $include{$FILE} = 'yes';
                  last;
               }
               else {
                  $include{$FILE} = 'no';
               }
            }
            elsif ($FORM{'case'} eq 'Sensitive') {
               if ($string =~ /$term/) {
		  $include{$FILE} = 'yes';
                  last;
               }
               else {
                  $include{$FILE} = 'no';
               }
            }
         }
      }
      if ($string =~ /<title>(.*)<\/title>/i) {
         $titles{$FILE} = "$1";
      }
      else {
         $titles{$FILE} = "$FILE";
      }
   }
}
      
sub return_html {
   print "Content-type: text/html\n\n";
   print "<html>\n <head>\n  <title>Elucid Software - Search Results</title>\n <style fprolloverstyle>A:hover \{color: #008FE6\} </style> \n </head>\n";
   print "<body topmargin=\"0\" leftmargin=\"0\" link=\"#003E62\" vlink=\"#003E62\" alink=\"#008FE6\">";
   print "<p><img border=\"0\" src=\"../../images/search.jpg\" width=\"403\" height=\"156\"></p>\n";
      
   print "<p style=\"margin-left: 15\; margin-right: 0\; margin-top: 0\; margin-bottom: 0\"><b><font face=\"Verdana\" size=\"2\">Search Results</font></b><br>\n";
   print "<font face=\"Verdana\" size=\"2\">Here ";
   print "are the result from the search you submitted.&nbsp\; Keep in mind these results ";
   print "may not be exact results.&nbsp\; The engine we use does a simple key word search ";
   print "on every page we have on our server and returns those pages as results.</font></p>";
   print "<p style=\"margin-left: 15\; margin-right: 0\; margin-top: 0\; margin-bottom: 0\">&nbsp\;</p>\n";

   print "<table border=\"0\" width=\"100%\">\n";
   print "<tr>\n";
   print "<td width=\"24%\" valign=\"top\">\n"; 
   
  
   print "<p style=\"margin: 0\"><font face=\"Tahoma\" size=\"2\"><b>Search Keys</b></font></p>\n";
   print "<p style=\"margin-left: 15\; margin-right: 0\; margin-top: 0\; margin-bottom: 0\">\n";

       $i = 0;
       foreach $term (@terms) {
          print "$term";
          $i++;
          if (!($i == @terms)) {
             print "&nbsp\;";
          }
       }
    
    
    print "</p>\n"; 
    print "<br>\n<br>\n";
    
    print "<form method=\"POST\" action=\"search.cgi\" name=\"search\">\n";
    print "<p><b><font face=\"Tahoma\" size=\"2\">Search:</font></b><br>\n";
    print "<input type=\"text\" name=\"terms\" size=\"24\"> <input type=\"submit\" value=\"GO\" style=\"font-family: Arial\; font-size: 10pt\"><br>\n";
    print "<font face=\"Tahoma\" size=\"1\"><a href=\"advancedsearch.htm\">Advanced Search</a></font></p>\n";
    print "<input type=\"hidden\" name=\"boolean\" value=\"AND\"><input type=\"hidden\" name=\"case\" value=\"Insensitive\">\n";
    print "</form>\n";

        
    print "</td>\n";
    
    print "<td width=\"2%\" valign=\"top\"><img border=\"0\" src=\"../../images/verticalline.gif\" width=\"1\" height=\"480\"></td>\n";
    print "<td width=\"74%\"valign=\"top\">\n";
    
   

   print "<p style=\"margin-top: 0\; margin-bottom: 0\"><font face=\"Verdana\" size=\"2\"><b>Search\n";
   print "Results</b></font></p>\n";
   

   foreach $key (keys %include) {
         if ($include{$key} eq 'yes') {
         print "<p style=\"margin-left: 15\; margin-top: 0\; margin-bottom: 0\"><font face=\"Verdana\" size=\"2\"><a href=\"$baseurl$key\">$titles{$key}</a></p>\n";
         }
   } 
   
   print "<br>\n";
   print "<hr size=\"1\" color=\"#003E62\">\n";
   print "<br>\n<br>\n";
   print "<p style=\"margin-left: 15\; margin-top: 0\; margin-bottom: 0\"><font size=\"2\" face=\"Verdana\"><a href=\"../../search.htm\">Search Page</a></font></p>\n";
      print "<p style=\"margin-left: 15\; margin-top: 0\; margin-bottom: 0\"><font size=\"2\" face=\"Verdana\"><a href=\"../../index.htm\">Main Page</a></font></p>\n";
   print "</td>\n";
   print "</tr>\n";
   print "</table>\n";
   print "<br>\n";
   print "<hr size=\"1\" color=\"#003E62\">\n";
   print "<br>\n";
   print "<p style=\"margin-top: 0\; margin-bottom: 0\"><font face=\"Verdana\" size=\"2\"><a href=\"http://www.worldwidemart.com/scripts/\">Script Written By Matt</a></font></p>\n";
   print "<p style=\"margin-top: 0\; margin-bottom: 0\"><font face=\"Verdana\" size=\"2\"><a href=\"http://elucidsoftware.hypermart.net\">Modified\n";
   print "By Elucid Software</a></font></p>\n";

   print "<br>\n";
   print "<!--#echo banner=\"\"-->\n";
   print "</body>\n</html>\n";
}
   
