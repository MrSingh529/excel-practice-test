import streamlit as st
import pandas as pd
import json
import datetime
from pathlib import Path
import hashlib
import plotly.express as px
import plotly.graph_objects as go

# Configure page
st.set_page_config(
    page_title="Excel Practice Test",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        text-align: center;
        color: #0078d4;
        border-bottom: 3px solid #0078d4;
        padding-bottom: 20px;
        margin-bottom: 30px;
    }
    .question-box {
        background: #f8f9fa;
        padding: 15px;
        border-left: 3px solid #28a745;
        border-radius: 4px;
        margin: 15px 0;
    }
    .instructions-box {
        background: #fff3cd;
        border: 1px solid #ffeaa7;
        padding: 15px;
        border-radius: 4px;
        margin: 20px 0;
    }
    .data-table {
        font-size: 12px;
    }
    .metric-card {
        background: white;
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'user_answers' not in st.session_state:
    st.session_state.user_answers = {}
if 'user_info' not in st.session_state:
    st.session_state.user_info = {}
if 'test_submitted' not in st.session_state:
    st.session_state.test_submitted = False

# Employee data
employee_data = [
    {"Employee": "Saravana Kumar R", "Gender": "Male", "Marital Status": "Unmarried", "Region": "South", "Location": "Trichy", "Department": "TSG & IT Hardware", "Total Amount Due": 2000},
    {"Employee": "Narsi Ram Meena", "Gender": "Male", "Marital Status": "Married", "Region": "North", "Location": "Lucknow", "Department": "TSG & IT Hardware", "Total Amount Due": 6000},
    {"Employee": "Shahbaz Khan", "Gender": "Male", "Marital Status": "Married", "Region": "North", "Location": "Agra", "Department": "Customer Service Division", "Total Amount Due": 4400},
    {"Employee": "Aman Mishra", "Gender": "Male", "Marital Status": "Unmarried", "Region": "West", "Location": "Satara", "Department": "TSG & IT Hardware", "Total Amount Due": 2300},
    {"Employee": "Bherulal Sharma", "Gender": "Male", "Marital Status": "Unmarried", "Region": "West", "Location": "Satara", "Department": "Accounts", "Total Amount Due": 10000},
    {"Employee": "Brajesh Sharma", "Gender": "Male", "Marital Status": "Married", "Region": "North", "Location": "Lucknow", "Department": "TSG & IT Hardware", "Total Amount Due": 15000},
    {"Employee": "Suraj Mahor", "Gender": "Male", "Marital Status": "Unmarried", "Region": "North", "Location": "Lucknow", "Department": "TSG & IT Hardware", "Total Amount Due": 14000},
    {"Employee": "Shikha Yadav", "Gender": "Female", "Marital Status": "Married", "Region": "East", "Location": "Noida", "Department": "Sales", "Total Amount Due": 200},
    {"Employee": "Sunita Gautam Dudhe", "Gender": "Female", "Marital Status": "Married", "Region": "West", "Location": "Nagpur", "Department": "Customer Service Division", "Total Amount Due": 123},
    {"Employee": "Dhan Das", "Gender": "Male", "Marital Status": "Unmarried", "Region": "East", "Location": "Guwahati", "Department": "TSG & IT Hardware", "Total Amount Due": 0},
    {"Employee": "Anamika Singh Chaudhary", "Gender": "Female", "Marital Status": "Unmarried", "Region": "North", "Location": "Agra", "Department": "Customer Service Division", "Total Amount Due": 1},
    {"Employee": "Chaitram Dhanraj Shahu", "Gender": "Male", "Marital Status": "Married", "Region": "West", "Location": "Nagpur", "Department": "Customer Service Division", "Total Amount Due": 12300},
    {"Employee": "Dev Singh Saharawat", "Gender": "Male", "Marital Status": "Unmarried", "Region": "North", "Location": "Ambala", "Department": "Accounts", "Total Amount Due": 5300},
    {"Employee": "Santosh Kumar Singh", "Gender": "Male", "Marital Status": "Married", "Region": "North", "Location": "Noida", "Department": "TSG & IT Hardware", "Total Amount Due": 47000},
]

# Correct answers
correct_answers = {
    "q1": "a",  # True
    "q2": "b",  # Column heading
    "q3": "b",  # Conditional Formatting
    "q4": "a",  # Right
    "q5": "b",  # Does not change
    "q6": "b",  # IF
    "q7": "a",  # True
    "q8": "a",  # Show only rows where Category = "Food"
}

# File paths for data storage
SUBMISSIONS_FILE = "test_submissions.json"
ADMIN_PASSWORD = "admin123"  # In production, use environment variables

def load_submissions():
    """Load existing submissions from file"""
    try:
        if Path(SUBMISSIONS_FILE).exists():
            with open(SUBMISSIONS_FILE, 'r') as f:
                return json.load(f)
    except:
        pass
    return []

def save_submission(submission):
    """Save a new submission"""
    submissions = load_submissions()
    submissions.append(submission)
    with open(SUBMISSIONS_FILE, 'w') as f:
        json.dump(submissions, f, indent=2)

def calculate_score(user_answers):
    """Calculate test score"""
    score = 0
    total = len(correct_answers)
    for q_id, correct_answer in correct_answers.items():
        if user_answers.get(q_id) == correct_answer:
            score += 1
    return score, total

def create_pivot_analysis():
    """Create pivot table analysis from employee data"""
    df = pd.DataFrame(employee_data)
    
    # Regional totals
    regional_totals = df.groupby('Region')['Total Amount Due'].sum().reset_index()
    
    # Department totals
    dept_totals = df.groupby('Department')['Total Amount Due'].sum().reset_index()
    
    # Regional gender count
    regional_gender = df.groupby(['Region', 'Gender']).size().reset_index(name='Count')
    
    return regional_totals, dept_totals, regional_gender

# Sidebar navigation
st.sidebar.title("Navigation")
page = st.sidebar.selectbox("Choose a page:", 
    ["🏠 Home", "📝 Take Test", "📊 Data Analysis", "👨‍💼 Admin Dashboard"])

if page == "🏠 Home":
    st.markdown('<h1 class="main-header">📊 Excel Practice Test</h1>', unsafe_allow_html=True)
    st.markdown('<p style="text-align: center; color: #666; font-style: italic;">Learning & Development Department: Together we learn, together we soar.</p>', unsafe_allow_html=True)
    
    st.markdown("""
    ## Welcome to the Digital Excel Practice Test!
    
    This comprehensive test evaluates your Excel knowledge across multiple areas:
    
    ### 📋 Test Sections:
    - **Section A**: Multiple Choice Questions (8 questions)
    - **Section B**: Data Analysis using Employee Dataset  
    - **Section C**: PivotTable Understanding
    
    ### 🎯 Learning Objectives:
    - Master Excel fundamentals
    - Understand data analysis concepts
    - Learn PivotTable functionality
    - Practice conditional formatting
    - Explore logical functions
    
    ### 📊 Features:
    - Interactive online test
    - Instant score calculation
    - Data visualization
    - Progress tracking
    - Admin monitoring
    """)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.info("**Total Questions**: 8 Multiple Choice")
    with col2:
        st.success("**Time Limit**: Self-paced")
    with col3:
        st.warning("**Passing Score**: 70%")

elif page == "📝 Take Test":
    st.markdown('<h1 class="main-header">📝 Excel Practice Test</h1>', unsafe_allow_html=True)
    
    if not st.session_state.test_submitted:
        # User information form
        with st.expander("📋 Enter Your Information", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                name = st.text_input("👤 Full Name*", value=st.session_state.user_info.get('name', ''))
                employee_id = st.text_input("🆔 Employee ID*", value=st.session_state.user_info.get('employee_id', ''))
            with col2:
                department = st.selectbox("🏢 Department*", 
                    ["", "TSG & IT Hardware", "Customer Service Division", "Accounts", "Sales", "HR", "Other"],
                    index=0 if not st.session_state.user_info.get('department') else 
                    ["", "TSG & IT Hardware", "Customer Service Division", "Accounts", "Sales", "HR", "Other"].index(st.session_state.user_info.get('department')))
                email = st.text_input("📧 Email*", value=st.session_state.user_info.get('email', ''))
        
        # Store user info
        st.session_state.user_info = {
            'name': name,
            'employee_id': employee_id,
            'department': department,
            'email': email
        }
        
        # Instructions
        st.markdown("""
        <div class="instructions-box">
        <strong>📝 Instructions:</strong><br>
        • Answer all 8 multiple-choice questions<br>
        • Select the best answer for each question<br>
        • Review the employee data table for context<br>
        • Submit your answers when complete<br>
        • You can change answers before final submission
        </div>
        """, unsafe_allow_html=True)
        
        # Questions
        st.markdown("## Section A: Multiple Choice Questions")
        
        questions = [
            {
                "id": "q1",
                "text": "The AutoSum feature adds up the numbers in a column or row that you specify.",
                "options": ["True", "False"]
            },
            {
                "id": "q2", 
                "text": 'In Excel, the label "AAA" (as seen at the top of a worksheet) is an example of a:',
                "options": ["Cell reference", "Column heading", "Name box entry", "Row heading"]
            },
            {
                "id": "q3",
                "text": "__________ quickly highlight important information in a spreadsheet by applying formatting options such as data bars, color scales, or icon sets.",
                "options": ["Cell references", "Conditional Formatting", "Excel tables", "PivotTables"]
            },
            {
                "id": "q4",
                "text": "As a rule, Excel will __________-align numbers in a cell.",
                "options": ["Right", "Left", "Top", "Bottom"]
            },
            {
                "id": "q5",
                "text": "When you copy a formula that contains an absolute reference (e.g., $A$1) to a new location, the absolute reference:",
                "options": ["Updates automatically to reflect the new row/column", "Does not change", "Becomes bold", "Gets a dotted outline in its cell"]
            },
            {
                "id": "q6",
                "text": "Which of the following is a logical function in Excel?",
                "options": ["AVERAGE", "IF", "SUMPRODUCT", "VLOOKUP"]
            },
            {
                "id": "q7",
                "text": "PivotTables are a powerful tool used to quickly group, summarize, and rearrange larger datasets.",
                "options": ["True", "False"]
            },
            {
                "id": "q8",
                "text": 'Consider a PivotTable with a slicer connected to the "Category" field. If you click "Food" on that slicer, the PivotTable will:',
                "options": ['Show only rows where Category = "Food"', 'Show all rows except those where Category = "Food"', 'Not change (slicer has no effect)']
            }
        ]
        
        for i, question in enumerate(questions, 1):
            st.markdown(f"""
            <div class="question-box">
            <strong>Question {i}:</strong> {question['text']}
            </div>
            """, unsafe_allow_html=True)
            
            option_labels = [chr(97 + j) for j in range(len(question['options']))]  # a, b, c, d
            formatted_options = [f"{label}. {option}" for label, option in zip(option_labels, question['options'])]
            
            selected = st.radio(
                f"Select your answer for Question {i}:",
                options=option_labels,
                format_func=lambda x: formatted_options[ord(x) - 97],
                key=question['id'],
                index=None
            )
            
            if selected:
                st.session_state.user_answers[question['id']] = selected
        
        # Employee Data Display
        st.markdown("## Section B: Employee Data Reference")
        st.markdown("*Use this data to understand the context for the questions above:*")
        
        df = pd.DataFrame(employee_data)
        st.dataframe(df, use_container_width=True)
        
        # Submit button
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("🚀 Submit Test", type="primary", use_container_width=True):
                # Validate user info
                if not all([name, employee_id, department, email]):
                    st.error("⚠️ Please fill in all required information fields!")
                elif len(st.session_state.user_answers) < len(questions):
                    st.error(f"⚠️ Please answer all questions! You have answered {len(st.session_state.user_answers)} out of {len(questions)} questions.")
                else:
                    # Calculate score
                    score, total = calculate_score(st.session_state.user_answers)
                    percentage = (score / total) * 100
                    
                    # Create submission record
                    submission = {
                        "timestamp": datetime.datetime.now().isoformat(),
                        "user_info": st.session_state.user_info,
                        "answers": st.session_state.user_answers,
                        "score": score,
                        "total": total,
                        "percentage": percentage
                    }
                    
                    # Save submission
                    save_submission(submission)
                    st.session_state.test_submitted = True
                    st.rerun()
    
    else:
        # Show results
        score, total = calculate_score(st.session_state.user_answers)
        percentage = (score / total) * 100
        
        st.success("🎉 Test Submitted Successfully!")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Score", f"{score}/{total}")
        with col2:
            st.metric("Percentage", f"{percentage:.1f}%")
        with col3:
            status = "PASS" if percentage >= 70 else "NEEDS IMPROVEMENT"
            color = "green" if percentage >= 70 else "red"
            st.metric("Result", status)
        
        # Detailed results
        st.markdown("## 📊 Detailed Results")
        results_data = []
        for i, (q_id, correct_answer) in enumerate(correct_answers.items(), 1):
            user_answer = st.session_state.user_answers.get(q_id, "Not answered")
            is_correct = user_answer == correct_answer
            results_data.append({
                "Question": i,
                "Your Answer": user_answer.upper() if user_answer != "Not answered" else user_answer,
                "Correct Answer": correct_answer.upper(),
                "Result": "✅ Correct" if is_correct else "❌ Incorrect"
            })
        
        results_df = pd.DataFrame(results_data)
        st.dataframe(results_df, use_container_width=True)
        
        if st.button("🔄 Take Test Again"):
            st.session_state.user_answers = {}
            st.session_state.test_submitted = False
            st.rerun()

elif page == "📊 Data Analysis":
    st.markdown('<h1 class="main-header">📊 Employee Data Analysis</h1>', unsafe_allow_html=True)
    
    df = pd.DataFrame(employee_data)
    regional_totals, dept_totals, regional_gender = create_pivot_analysis()
    
    # Overview metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Employees", len(df))
    with col2:
        st.metric("Total Amount Due", f"₹{df['Total Amount Due'].sum():,}")
    with col3:
        st.metric("Regions", df['Region'].nunique())
    with col4:
        st.metric("Departments", df['Department'].nunique())
    
    # Charts
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("💰 Regional Totals")
        fig1 = px.bar(regional_totals, x='Region', y='Total Amount Due', 
                     title="Total Amount Due by Region",
                     color='Total Amount Due',
                     color_continuous_scale='Blues')
        st.plotly_chart(fig1, use_container_width=True)
        
        st.subheader("👥 Gender Distribution by Region")
        fig3 = px.bar(regional_gender, x='Region', y='Count', color='Gender',
                     title="Employee Count by Region and Gender",
                     barmode='group')
        st.plotly_chart(fig3, use_container_width=True)
    
    with col2:
        st.subheader("🏢 Department Totals")
        fig2 = px.pie(dept_totals, values='Total Amount Due', names='Department',
                     title="Total Amount Due by Department")
        st.plotly_chart(fig2, use_container_width=True)
        
        st.subheader("📍 Location Distribution")
        location_counts = df['Location'].value_counts()
        fig4 = px.bar(x=location_counts.index, y=location_counts.values,
                     title="Employee Count by Location")
        st.plotly_chart(fig4, use_container_width=True)
    
    # Data tables
    st.subheader("📋 Pivot Table Results")
    
    tab1, tab2, tab3 = st.tabs(["Regional Totals", "Department Totals", "Regional Gender Count"])
    
    with tab1:
        st.dataframe(regional_totals, use_container_width=True)
    
    with tab2:
        st.dataframe(dept_totals.sort_values('Total Amount Due', ascending=False), use_container_width=True)
    
    with tab3:
        pivot_gender = regional_gender.pivot(index='Region', columns='Gender', values='Count').fillna(0)
        st.dataframe(pivot_gender, use_container_width=True)

elif page == "👨‍💼 Admin Dashboard":
    st.markdown('<h1 class="main-header">👨‍💼 Admin Dashboard</h1>', unsafe_allow_html=True)
    
    # Admin authentication
    if 'admin_authenticated' not in st.session_state:
        st.session_state.admin_authenticated = False
    
    if not st.session_state.admin_authenticated:
        st.warning("🔐 Admin access required")
        password = st.text_input("Enter admin password:", type="password")
        if st.button("Login"):
            if password == ADMIN_PASSWORD:
                st.session_state.admin_authenticated = True
                st.success("✅ Admin access granted!")
                st.rerun()
            else:
                st.error("❌ Invalid password!")
    else:
        submissions = load_submissions()
        
        if not submissions:
            st.info("📝 No test submissions yet.")
        else:
            # Summary statistics
            total_submissions = len(submissions)
            avg_score = sum(s['percentage'] for s in submissions) / total_submissions
            pass_rate = sum(1 for s in submissions if s['percentage'] >= 70) / total_submissions * 100
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Submissions", total_submissions)
            with col2:
                st.metric("Average Score", f"{avg_score:.1f}%")
            with col3:
                st.metric("Pass Rate", f"{pass_rate:.1f}%")
            
            # Charts
            col1, col2 = st.columns(2)
            
            with col1:
                # Score distribution
                scores = [s['percentage'] for s in submissions]
                fig1 = px.histogram(x=scores, title="Score Distribution", 
                                  nbins=10, labels={'x': 'Score %', 'y': 'Count'})
                fig1.add_vline(x=70, line_dash="dash", line_color="red", 
                              annotation_text="Pass Line (70%)")
                st.plotly_chart(fig1, use_container_width=True)
            
            with col2:
                # Department wise performance
                dept_scores = {}
                for s in submissions:
                    dept = s['user_info']['department']
                    if dept not in dept_scores:
                        dept_scores[dept] = []
                    dept_scores[dept].append(s['percentage'])
                
                dept_avg = {dept: sum(scores)/len(scores) for dept, scores in dept_scores.items()}
                fig2 = px.bar(x=list(dept_avg.keys()), y=list(dept_avg.values()),
                             title="Average Score by Department",
                             labels={'x': 'Department', 'y': 'Average Score %'})
                st.plotly_chart(fig2, use_container_width=True)
            
            # Detailed submissions table
            st.subheader("📋 All Submissions")
            
            # Prepare data for display
            display_data = []
            for s in submissions:
                display_data.append({
                    "Timestamp": s['timestamp'][:19].replace('T', ' '),
                    "Name": s['user_info']['name'],
                    "Employee ID": s['user_info']['employee_id'],
                    "Department": s['user_info']['department'],
                    "Email": s['user_info']['email'],
                    "Score": f"{s['score']}/{s['total']}",
                    "Percentage": f"{s['percentage']:.1f}%",
                    "Status": "PASS" if s['percentage'] >= 70 else "FAIL"
                })
            
            submissions_df = pd.DataFrame(display_data)
            st.dataframe(submissions_df, use_container_width=True)
            
            # Download submissions
            if st.button("📥 Download All Submissions (JSON)"):
                st.download_button(
                    label="Download Submissions",
                    data=json.dumps(submissions, indent=2),
                    file_name=f"excel_test_submissions_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                    mime="application/json"
                )
        
        if st.button("🚪 Admin Logout"):
            st.session_state.admin_authenticated = False
            st.rerun()

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #666;'>"
    "📊 Excel Practice Test | Learning & Development Department<br>"
    "Together we learn, together we soar 🚀"
    "</div>", 
    unsafe_allow_html=True
)