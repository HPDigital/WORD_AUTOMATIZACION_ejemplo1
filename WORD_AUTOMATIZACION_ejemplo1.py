"""
WORD_AUTOMATIZACION_ejemplo1
"""

#!/usr/bin/env python
# coding: utf-8

# In[11]:


import os, sys
from docxtpl import DocxTemplate as DocTemp
from docxtpl import InlineImage
from docx.shared import Cm, Inches, Mm, Emu


# In[12]:


path_in = r"C:\Users\HP\Desktop\Propuesta DIPLOMADO\Temaplate_prueba\temp_prueba.docx"
path_out = r"C:\Users\HP\Desktop\Propuesta DIPLOMADO\Temaplate_prueba\temp_prueba1.docx"

img_in = r"C:\Users\HP\Pictures\Saved Pictures\VALUE_MAP.png"


# In[13]:


doc = DocTemp(path_in)
imagen = InlineImage(doc, img_in, Cm(5))
context = {'titulo': 'titulango', 'resumen':'resumensingo', 'imagen': imagen}
doc.render(context)
doc.save(path_out)


# In[ ]:




