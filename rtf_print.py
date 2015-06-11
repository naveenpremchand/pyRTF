from PyRTF import *
import os
import re

class PrintingBase:
    def __init__(self,):
        self.lst_column_title =  []
        self.lst_header_part = []
        self.lst_data = []
        self.lst_footer_data = []
        self.str_title = ''
        pass
    
class RTF:
    
        def __init__(self,str_file_name):
                self.str_file_name = str_file_name
                self.init_document()
                pass

	def init_document(self,):
		self.rtf_document = Document()
		self.rtf_doc_style = self.rtf_document.StyleSheet
		pass

	def add_new_section(self,):
		self.section = Section()
		pass

        def create_new_table(self,lst_col_info):
                self.table = Table(*lst_col_info)
                pass

        def display_page_header(self,str_title):
            self.add_new_section()
            para = Paragraph( self.rtf_doc_style.ParagraphStyles.Heading1 ,ParagraphPS(alignment = 3))
            para.append( Text(str_title,TextPS( colour = self.rtf_doc_style.Colours.Red )) )
            self.section.append( para )
            self.rtf_document.Sections.append(self.section)
            pass
        

        def draw_header_footer_notes(self,lst_header = []):
                lst_col_info = [TabPS.DEFAULT_WIDTH * 5,TabPS.DEFAULT_WIDTH * 4,TabPS.DEFAULT_WIDTH * 4]
                self.create_new_table(lst_col_info)
        
                for tpl_header in lst_header:
                    lst_para = []
                    for str_text in tpl_header:
                        
                        if str_text.find(':') <> -1:
                            # Need to match any bold characters
                            match = re.search(r'\s*([\w.-]+)+\s*(:)\s*%s\s*([\w.-]+)+\s*%s'%('<b>','</b>'),str_text)
                            if match:
                                tps = TextPS(bold = True)
                                text = Text( match.group(3), tps )
                                int_align = 1
                                if match.group(2).isdigit():
                                   int_align = 2
                                   pass
                               
                                ps = ParagraphPS(alignment = int_align)
                                lst_para.append(Paragraph(match.group(1),' : ',text,ps,self.rtf_doc_style.ParagraphStyles.Normal))

                            else:
                                match = re.search(r'%s\s*([\w.-]+)+\s*%s'%('<b>','</b>'),str_text)
                                int_align = 1
                                if match:
                                    if match.group(1).isdigit():
                                        int_align = 2
                                    pass

                                ps = ParagraphPS(alignment = int_align)
                                lst_para.append(Paragraph(str_text,ps,self.rtf_doc_style.ParagraphStyles.Normal))
                                pass

                        else:
                            match = re.search(r'\s*([\w.-]+)+\s*%s\s*([\w.-]+)+\s*%s'%('<b>','</b>'),str_text)
                            
                            if match:
                                tps = TextPS(bold = True)
                                int_align = 1
                                if match.group(2).isdigit():
                                   int_align = 2
                                   pass

                                ps = ParagraphPS(alignment = int_align)
                            
                                text = Text( match.group(2), tps )
                                lst_para.append(Paragraph(match.group(1),' : ',text,ps,self.rtf_doc_style.ParagraphStyles.Normal))
                            else:
                                
                                match = re.search(r'%s\s*([\w.-]+)+\s*%s'%('<b>','</b>'),str_text)
                                
                                int_align = 1
                                
                                if match and match.group(1).isdigit():
                                   int_align = 2
                                   pass

                                lst_para.append(Paragraph(str_text, self.rtf_doc_style.ParagraphStyles.Normal))
                            pass
                        pass
                    
                    if len(lst_para) == 3:
                        c1 = Cell(lst_para[0])
                        c2 = Cell(lst_para[1])
                        c3 = Cell(lst_para[2])
                        self.table.AddRow( c1, c2, c3 )

                self.section.append(self.table)
                pass

        def create_space_before(self,int_line):
            para_props = ParagraphPS(space_before = int_line)
            para = Paragraph(para_props)
            self.section.append(para)
            pass

        def draw_column_heading(self,lst_column_title):
            dct_column_width = self.compute_column_width(lst_column_title)

            self.create_space_before(1000)

            lst_col_width = []
            int_max_width = 13
            
            for int_col_index in dct_column_width:
                tpl_col_width = dct_column_width[int_col_index]
                # Check whether column is hidden or not
                if tpl_col_width[1][3]:
                    int_column_width = TabPS.DEFAULT_WIDTH * (tpl_col_width[0]*int_max_width)
                    lst_col_width.append(int(int_column_width))
                pass

            self.create_new_table(lst_col_width)
            

            lst_column_names = []
            int_index = 1
            for tpl_header in lst_column_title:

                tpl_col_width = dct_column_width[int_index]

                if tpl_col_width[1][3]:
                    # Need shading colour for column heading
                    sps = ShadingPS(pattern = 8 ,foreground = self.rtf_doc_style.Colours.Turquoise)
                    str_column_name = Paragraph(Text(tpl_header[1],TextPS(bold = True)),sps)
                    lst_column_names.append(Cell(str_column_name))
                    pass
                
                int_index += 1

        
            self.table.AddRow(*lst_column_names)

            pass
        
        def compute_column_width(self,lst_column_title):
            # Need to compute the actual column width from given
            dct_column_width   = {}
            int_max_len = sum([tpl_column[2] for tpl_column in lst_column_title if tpl_column[3]])
            for tpl_column in lst_column_title:
                dct_column_width[tpl_column[0]] = (float(tpl_column[2])/int_max_len,tpl_column)
                pass
            
            return dct_column_width

        def draw_column_data(self,lst_column_data,lst_column_title):
            dct_column_prop = self.compute_column_width(lst_column_title)
            
            
            for tpl_column in lst_column_data:
                lst_col_data = []
                int_index = 1
                for str_column in tpl_column:
                    tpl_col_prop = dct_column_prop[int_index][1]
                    int_align = 1
                    # Check whether it is a number or not
                    if tpl_col_prop[4]:
                        int_align = 2
                        pass
                    ps = ParagraphPS(alignment = int_align)
                    # Check whether column is hidden or not
                    if tpl_col_prop[3]:
                        lst_col_data.append(Cell(Paragraph(str(str_column),ps)))
                    int_index += 1

                self.table.AddRow(*lst_col_data)
                
                pass
            
            pass


        def draw_data(self,ins_printing_base):

                self.display_page_header(ins_printing_base.str_title)
                self.create_space_before(1000)
                self.add_new_section()

                self.rtf_document.Sections.append(self.section)

               
                # Draw header notes
                self.draw_header_footer_notes(ins_printing_base.lst_header_part)
                
                
                # Draw column heading
                self.draw_column_heading(ins_printing_base.lst_column_title)

              
                # Draw column data
                self.draw_column_data(ins_printing_base.lst_data,ins_printing_base.lst_column_title)

                self.section.append(self.table)

               
                self.create_space_before(1000)

                
                # Draw footer data
                self.draw_header_footer_notes(ins_printing_base.lst_footer_data)
                
                return self.rtf_document
                pass

	def open_writable_file(self,):
		str_file_path = os.path.join(os.environ['HOME'],'Desktop','RTF',self.str_file_name)
		return file( '%s.rtf' % str_file_path, 'w' )
		pass
            


if __name__ == '__main__' :
    
    ins_render = Renderer()

    str_file_name = 'sample'

    ins_rtf = RTF(str_file_name)

    str_output_file = ins_rtf.open_writable_file()

    ins_printing_base = PrintingBase()

    ins_printing_base.str_title = 'Purchase Voucher'
    # This is the header part each row data is included in each tuple
    ins_printing_base.lst_header_part = [('Holder Name :<b>Sam </b>','Date:<b> 22-12-2015 </b>','Mode : Debit card ')]

    # This is a column title 
    # each tuple follows order '1' : Column index (column no), 'SL.No': Column Name,'2': Column width,True : Display/hide , True : is a  number field
    ins_printing_base.lst_column_title = [(1,'SL.No',2,True,True),(2,'Product Code',4,True,False),(3,'Product Name',4,True,False),
                                   (4,'Quantity',3,True,True),(5,'Price',3,True,True),(6,'Total',3,True,True)]

    
    # Data for printing based on column title
    ins_printing_base.lst_data = [(1,'M/A','Mobile & Accessories ',4,50,200),(2,'C/P','Computer & Parts',4,10,40)]

    # Footer Data for printing 
    ins_printing_base.lst_footer_data = [('Two Hundred and forty Rupees','','Total  <b>240</b>')]

    rtf_doc = ins_rtf.draw_data(ins_printing_base)

    ins_render.Write( rtf_doc, str_output_file)
    
    pass

