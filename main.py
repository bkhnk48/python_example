import tkinter  # Import thư viện Tkinter để tạo giao diện người dùng
from tkinter import ttk  # ttk là một module của Tkinter, cung cấp một số widget mở rộng
from tkinter import messagebox  # messagebox là một module của Tkinter, dùng để hiển thị các hộp thoại thông báo
import os  # Import thư viện os để tương tác với hệ điều hành
import openpyxl  # Import thư viện openpyxl để làm việc với file Excel

# Định nghĩa hàm enter_data để xử lý dữ liệu nhập vào từ form
def enter_data():
    accepted = accept_var.get()  # Lấy giá trị của biến accept_var
    
    if accepted=="Accepted":  # Nếu người dùng đã chấp nhận các điều khoản
        # Lấy thông tin người dùng
        firstname = first_name_entry.get()  # Lấy tên từ entry tên
        lastname = last_name_entry.get()  # Lấy họ từ entry họ
        
        if firstname and lastname:  # Nếu cả tên và họ đều không rỗng
            title = title_combobox.get()  # Lấy danh xưng từ combobox danh xưng
            age = age_spinbox.get()  # Lấy tuổi từ spinbox tuổi
            nationality = nationality_combobox.get()  # Lấy quốc tịch từ combobox quốc tịch
            
            # Lấy thông tin khóa học
            registration_status = reg_status_var.get()  # Lấy trạng thái đăng ký từ biến reg_status_var
            numcourses = numcourses_spinbox.get()  # Lấy số lượng khóa học từ spinbox số lượng khóa học
            numsemesters = numsemesters_spinbox.get()  # Lấy số học kỳ từ spinbox số học kỳ
            
            # In thông tin ra console
            print("First name: ", firstname, "Last name: ", lastname)
            print("Title: ", title, "Age: ", age, "Nationality: ", nationality)
            print("# Courses: ", numcourses, "# Semesters: ", numsemesters)
            print("Registration status", registration_status)
            print("------------------------------------------")
            
            # Đường dẫn tới file Excel
            filepath = "D:\codefirst.io\Tkinter Data Entry\data.xlsx"
            
            # Nếu file Excel không tồn tại, tạo mới file và thêm tiêu đề cột
            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()  # Tạo workbook mới
                sheet = workbook.active  # Lấy sheet đang hoạt động
                heading = ["First Name", "Last Name", "Title", "Age", "Nationality",
                           "# Courses", "# Semesters", "Registration status"]  # Tiêu đề cột
                sheet.append(heading)  # Thêm tiêu đề vào sheet
                workbook.save(filepath)  # Lưu workbook
            
            # Mở file Excel và thêm dữ liệu
            workbook = openpyxl.load_workbook(filepath)  # Mở workbook
            sheet = workbook.active  # Lấy sheet đang hoạt động
            # Thêm dữ liệu vào sheet
            sheet.append([firstname, lastname, title, age, nationality, numcourses,
                          numsemesters, registration_status])
            workbook.save(filepath)  # Lưu workbook
                
        else:  # Nếu tên hoặc họ rỗng
            # Hiển thị hộp thoại cảnh báo
            tkinter.messagebox.showwarning(title="Error", message="First name and last name are required.")
    else:  # Nếu người dùng chưa chấp nhận các điều khoản
        # Hiển thị hộp thoại cảnh báo
        tkinter.messagebox.showwarning(title= "Error", message="You have not accepted the terms")
window = tkinter.Tk()  # Tạo một cửa sổ mới
window.title("Data Entry Form")  # Đặt tiêu đề cho cửa sổ

frame = tkinter.Frame(window)  # Tạo một frame mới trong cửa sổ
frame.pack()  # Đặt frame vào cửa sổ

# Tạo frame chứa thông tin người dùng
user_info_frame = tkinter.LabelFrame(frame, text="User Information")  # Tạo LabelFrame mới với tiêu đề là "User Information"
user_info_frame.grid(row=0, column=0, padx=20, pady=10)  # Đặt frame vào vị trí (0, 0) trên grid

# Tạo các label và entry cho tên và họ
first_name_label = tkinter.Label(user_info_frame, text="First Name")  # Tạo label cho tên
first_name_label.grid(row=0, column=0)  # Đặt label vào vị trí (0, 0) trên grid
last_name_label = tkinter.Label(user_info_frame, text="Last Name")  # Tạo label cho họ
last_name_label.grid(row=0, column=1)  # Đặt label vào vị trí (0, 1) trên grid

first_name_entry = tkinter.Entry(user_info_frame)  # Tạo entry cho tên
first_name_entry.grid(row=1, column=0)  # Đặt entry vào vị trí (1, 0) trên grid
last_name_entry = tkinter.Entry(user_info_frame)  # Tạo entry cho họ
last_name_entry.grid(row=1, column=1)  # Đặt entry vào vị trí (1, 1) trên grid

# Tạo label và combobox cho danh xưng
title_label = tkinter.Label(user_info_frame, text="Title")  # Tạo label cho danh xưng
title_combobox = ttk.Combobox(user_info_frame, values=["", "Mr.", "Ms.", "Dr."])  # Tạo combobox cho danh xưng
title_label.grid(row=0, column=2)  # Đặt label vào vị trí (0, 2) trên grid
title_combobox.grid(row=1, column=2)  # Đặt combobox vào vị trí (1, 2) trên grid

# Tạo label và spinbox cho tuổi
age_label = tkinter.Label(user_info_frame, text="Age")  # Tạo label cho tuổi
age_spinbox = tkinter.Spinbox(user_info_frame, from_=18, to=110)  # Tạo spinbox cho tuổi
age_label.grid(row=2, column=0)  # Đặt label vào vị trí (2, 0) trên grid
age_spinbox.grid(row=3, column=0)  # Đặt spinbox vào vị trí (3, 0) trên grid

# Tạo label và combobox cho quốc tịch
nationality_label = tkinter.Label(user_info_frame, text="Nationality")  # Tạo label cho quốc tịch
nationality_combobox = ttk.Combobox(user_info_frame, values=["Africa", "Antarctica", "Asia", "Europe", "North America", "Oceania", "South America"])  # Tạo combobox cho quốc tịch
nationality_label.grid(row=2, column=1)  # Đặt label vào vị trí (2, 1) trên grid
nationality_combobox.grid(row=3, column=1)  # Đặt combobox vào vị trí (3, 1) trên grid

# Đặt padding cho tất cả các widget con của user_info_frame
for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Tạo frame chứa thông tin khóa học
courses_frame = tkinter.LabelFrame(frame)  # Tạo LabelFrame mới
courses_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)  # Đặt frame vào vị trí (1, 0) trên grid

# Tạo label và checkbutton cho trạng thái đăng ký
registered_label = tkinter.Label(courses_frame, text="Registration Status")  # Tạo label cho trạng thái đăng ký
reg_status_var = tkinter.StringVar(value="Not Registered")  # Tạo biến StringVar với giá trị mặc định là "Not Registered"
registered_check = tkinter.Checkbutton(courses_frame, text="Currently Registered",  # Tạo Checkbutton
                                       variable=reg_status_var, onvalue="Registered", offvalue="Not registered")
registered_label.grid(row=0, column=0)  # Đặt label vào vị trí (0, 0) trên grid
registered_check.grid(row=1, column=0)  # Đặt Checkbutton vào vị trí (1, 0) trên grid
numcourses_label = tkinter.Label(courses_frame, text= "# Completed Courses")  # Tạo label cho số lượng khóa học đã hoàn thành
numcourses_spinbox = tkinter.Spinbox(courses_frame, from_=0, to='infinity')  # Tạo spinbox cho số lượng khóa học đã hoàn thành
numcourses_label.grid(row=0, column=1)  # Đặt label vào vị trí (0, 1) trên grid
numcourses_spinbox.grid(row=1, column=1)  # Đặt spinbox vào vị trí (1, 1) trên grid

numsemesters_label = tkinter.Label(courses_frame, text="# Semesters")  # Tạo label cho số học kỳ
numsemesters_spinbox = tkinter.Spinbox(courses_frame, from_=0, to="infinity")  # Tạo spinbox cho số học kỳ
numsemesters_label.grid(row=0, column=2)  # Đặt label vào vị trí (0, 2) trên grid
numsemesters_spinbox.grid(row=1, column=2)  # Đặt spinbox vào vị trí (1, 2) trên grid

for widget in courses_frame.winfo_children():  # Duyệt qua tất cả các widget con của courses_frame
    widget.grid_configure(padx=10, pady=5)  # Đặt padding cho mỗi widget

# Tạo frame cho điều khoản và điều kiện
terms_frame = tkinter.LabelFrame(frame, text="Terms & Conditions")  # Tạo LabelFrame mới với tiêu đề là "Terms & Conditions"
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)  # Đặt frame vào vị trí (2, 0) trên grid

# Tạo Checkbutton cho việc chấp nhận điều khoản và điều kiện
accept_var = tkinter.StringVar(value="Not Accepted")  # Tạo biến StringVar với giá trị mặc định là "Not Accepted"
terms_check = tkinter.Checkbutton(terms_frame, text= "I accept the terms and conditions.",
                                  variable=accept_var, onvalue="Accepted", offvalue="Not Accepted")  # Tạo Checkbutton
terms_check.grid(row=0, column=0)  # Đặt Checkbutton vào vị trí (0, 0) trên grid

# Tạo Button để nhập dữ liệu
button = tkinter.Button(frame, text="Enter data", command= enter_data)  # Tạo Button với text là "Enter data" và command là hàm enter_data
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)  # Đặt Button vào vị trí (3, 0) trên grid
 
window.mainloop()  # Bắt đầu vòng lặp chính của window
