from win32com.client import Dispatch

wdFormatPDF = 17


def doc_to_pdf(file_path, suffix):
    word = Dispatch('Word.Application')
    doc = word.Documents.Open(file_path)
    doc.SaveAs(file_path.replace(suffix, ".pdf"), FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()


if __name__ == "__main__":
    """
      doc_files = []
      directory = "C:\\Users\\localhost\\Desktop"
      doc_to_pdf(r"C:\\Users\\localhost\\Desktop\\aa.doc", ".doc")
      for root, dirs, filenames in walk(directory):
          for file in filenames:
              # print(root, dirs, file)
              print(join(root, file))
      print("_" * 120)
      for file in os.listdir(r'C:\\Users\\localhost\\Desktop'):
          if os.path.isfile(join(r'C:\\Users\\localhost\\Desktop', file)):
              print(file)
    """
    a = [["a", "b"], ["c", "d"], [1, ""], ["", True]]
    # print(int(a))
    # import numpy as np
    #
    # b = np.array(a)
    # print(b[:, 0])
    # print(b[:, 1])
    # print(type(b[2, 0]))
    a = [(1, '201'), (2, '219'), (4, '21901'), (5, '203'), (6, '204'), (7, '205')]
    b = ['5', '6']
    # print(a[8])
    try:
        print(a[8])
    except Exception as e:
        print(e)

