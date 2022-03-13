# import
from ResumeExtracter import ResumeExtracter
# main function
if __name__ == '__main__':
    resumes_folder_path = r'D:\Resume'
    resume_extracter = ResumeExtracter(resumes_folder_path)
    resume_extracter.to_excel('test.xlsx', resumes_folder_path)