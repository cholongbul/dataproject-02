import pandas as pd
from docx import Document

# document = Document("../NIA-IFT-OpenAPI활용가이드-03.국민안전처_소방항공기보유현황-V0.6 (1).docx")
# tables = document.tables
# for table in tables:
#     data = [[cell.text for cell in row.cells] for row in table.rows]
#     docdf = pd.DataFrame(data)
#     print(docdf)

import olefile
import pandas as pd
f = olefile.OleFileIO('../API청양통계.hwp')
#PrvText 스트림 내의 내용을 읽기
encoded_text = f.openstream('PrvText').read()
#인코딩된 텍스트를 UTF-16으로 디코딩
decoded_text = encoded_text.decode('UTF-16')
print(decoded_text)