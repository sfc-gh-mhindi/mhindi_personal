import pandas as pd
from datetime import datetime, timedelta
import math

def calculate_mortgage_schedule(
    loan_amount,
    payment_type,
    payment_amount,
    duration_years,
    annual_interest_rate,
    offset_start_balance,
    offset_growth_per_month,
    loan_start_date_str
):
    """
    Calculates the mortgage amortization schedule with an offset account and generates an Excel file.

    Args:
        loan_amount (float): The initial loan amount.
        payment_type (str): The frequency of payments ('Weekly', 'Fortnightly', 'Monthly').
        payment_amount (float): The regular payment amount per period.
        duration_years (int): The original duration of the loan in years.
        annual_interest_rate (float): The annual interest rate (e.g., 0.0649 for 6.49%).
        offset_start_balance (float): The initial balance in the offset account.
        offset_growth_per_month (float): The amount by which the offset account grows each month.
        loan_start_date_str (str): The start date of the loan (DD/MM/YYYY).

    Returns:
        pandas.DataFrame: A DataFrame containing the full amortization schedule.
    """

    # --- Input Processing and Rate Conversion ---
    
    # Convert annual interest rate to periodic rate
    if payment_type == "Weekly":
        periods_per_year = 52
    elif payment_type == "Fortnightly":
        periods_per_year = 26
    elif payment_type == "Monthly":
        periods_per_year = 12
    else:
        raise ValueError("Invalid payment_type. Choose 'Weekly', 'Fortnightly', or 'Monthly'.")

    periodic_interest_rate = (1 + annual_interest_rate)**(1/periods_per_year) - 1

    # Convert loan start date string to datetime object
    loan_start_date = datetime.strptime(loan_start_date_str, "%d/%m/%Y")

    # Calculate offset growth per period based on monthly growth
    offset_growth_per_period = offset_growth_per_month / (periods_per_year / 12)

    # --- Amortization Logic ---

    current_loan_balance = loan_amount
    current_offset_balance = offset_start_balance
    period_counter = 0
    schedule_data = []

    # Keep track of loan year and period in loan year
    loan_year_start_offset = (loan_start_date.month, loan_start_date.day)
    period_in_loan_year_counts = {}

    # Loop until loan is paid off (or maximum periods reached for safety)
    max_periods = duration_years * periods_per_year * 2 # Safety break: up to double the original term
    
    while current_loan_balance > 0 and period_counter < max_periods:
        period_counter += 1
        
        # Store starting balance for this period
        starting_loan_balance_this_period = current_loan_balance
        
        # Calculate current payment date
        current_payment_date = loan_start_date + timedelta(days=(period_counter - 1) * (365.25 / periods_per_year))
        # Adjust days slightly for more precise fortnight/week intervals if needed, but timedelta handles this generally well for fixed periods.
        if payment_type == "Fortnightly":
            current_payment_date = loan_start_date + timedelta(days=(period_counter - 1) * 14)
        elif payment_type == "Weekly":
            current_payment_date = loan_start_date + timedelta(days=(period_counter - 1) * 7)
        elif payment_type == "Monthly":
            # For monthly, more complex date arithmetic is needed to ensure end-of-month payments.
            # Simple timedelta might drift. For simplicity, assume fixed-day monthly (e.g., 15th of each month)
            # or adjust slightly. Given the previous example, this might be less critical.
            # A more robust monthly calculation would be:
            # next_month_date = current_payment_date.replace(day=1) + timedelta(days=32)
            # current_payment_date = next_month_date.replace(day=min(loan_start_date.day, next_month_date.day))
            # For this script, sticking to timedelta approximations for consistency with previous.
            pass # current_payment_date calculated above is sufficient for this simple model

        day_of_week = current_payment_date.strftime("%A")

        # Determine loan year and period within that loan year
        current_year = current_payment_date.year
        current_month = current_payment_date.month
        current_day = current_payment_date.day

        # Calculate loan_year based on loan_start_date
        if (current_month > loan_year_start_offset[0]) or \
           (current_month == loan_year_start_offset[0] and current_day >= loan_year_start_offset[1]):
            loan_year_val = current_year - loan_start_date.year + 1
        else:
            loan_year_val = current_year - loan_start_date.year
        
        # Calculate period_in_loan_year
        if loan_year_val not in period_in_loan_year_counts:
            period_in_loan_year_counts[loan_year_val] = 1
        else:
            period_in_loan_year_counts[loan_year_val] += 1
        
        period_in_loan_year_val = period_in_loan_year_counts[loan_year_val]

        # Calculate effective loan balance for interest calculation (cannot be negative)
        effective_loan_balance_for_interest = max(0, current_loan_balance - current_offset_balance)

        # Calculate interest paid for the period
        interest_paid = effective_loan_balance_for_interest * periodic_interest_rate

        # Calculate principal paid for the period
        # The principal paid is the total payment minus the interest paid
        principal_paid = payment_amount - interest_paid

        # Handle final payment: ensure it doesn't overpay the loan
        if current_loan_balance - principal_paid < 0:
            principal_paid = current_loan_balance
            actual_payment_this_period = principal_paid + interest_paid
            current_loan_balance = 0
        else:
            actual_payment_this_period = payment_amount
            current_loan_balance -= principal_paid

        schedule_data.append({
            f"{payment_type}": period_counter,
            "Year": loan_year_val,
            f"{payment_type} in Year": period_in_loan_year_val,
            "Payment Date (DD/MM/YYYY)": current_payment_date.strftime("%d/%m/%Y"),
            "Payment Date (Day of Week)": day_of_week,
            "Starting Loan Balance": starting_loan_balance_this_period,
            "Offset Account Balance": current_offset_balance,
            "Effective Loan Balance (for Interest)": effective_loan_balance_for_interest,
            "Interest Paid": interest_paid,
            "Principal Paid": principal_paid,
            f"{payment_type} Payment": actual_payment_this_period,
            "Remaining Loan Balance": current_loan_balance
        })
        # Note: starting_loan_balance_this_period is captured at the beginning of each iteration

        # Update offset balance for the next period
        current_offset_balance += offset_growth_per_period

        # Break if loan is paid off
        if current_loan_balance <= 0:
            break
            
    total_actual_periods = period_counter

    # --- Calculate Remaining Term (Years & Months) for each entry ---
    for i, row in enumerate(schedule_data):
        remaining_periods = total_actual_periods - (i + 1) # i is 0-indexed, period_counter is 1-indexed
        
        remaining_years = remaining_periods // periods_per_year
        remaining_months_fraction = (remaining_periods % periods_per_year) / periods_per_year * 12
        remaining_months = round(remaining_months_fraction)

        # Adjust months if it rounds up to 12
        if remaining_months == 12:
            remaining_years += 1
            remaining_months = 0

        row["Remaining Term (Years & Months)"] = f"{remaining_years} Years, {remaining_months} Months"

    # Create DataFrame
    df = pd.DataFrame(schedule_data)
    
    # Store unformatted data for calculations
    df_unformatted = df.copy()
    
    # Format currency columns with $ and 2 decimal places
    currency_columns = [
        "Starting Loan Balance", 
        "Offset Account Balance", 
        "Effective Loan Balance (for Interest)",
        "Interest Paid", 
        "Principal Paid", 
        f"{payment_type} Payment", 
        "Remaining Loan Balance"
    ]
    
    for col in currency_columns:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: f"${x:,.2f}")
    
    # Add unformatted data as an attribute for calculations
    # Use setattr to avoid pandas warning
    setattr(df, 'unformatted', df_unformatted)
    
    return df

if __name__ == "__main__":
    print("--- Mortgage Amortization Calculator with Offset ---")

    # Get inputs from the user with default values
    try:
        loan_amount_input = input("Enter Loan Amount [default: $759,476.79]: $").strip()
        loan_amount = float(loan_amount_input.replace(',', '')) if loan_amount_input else 759476.79
        
        payment_type_input = input("Enter Payment Type (Weekly, Fortnightly, or Monthly) [default: Fortnightly]: ").strip()
        payment_type = payment_type_input.capitalize() if payment_type_input else "Fortnightly"
        if payment_type not in ["Weekly", "Fortnightly", "Monthly"]:
            print("Invalid payment type. Please choose 'Weekly', 'Fortnightly', or 'Monthly'.")
            exit()

        payment_amount_input = input(f"Enter {payment_type} Payment [default: $2,410]: $").strip()
        payment_amount = float(payment_amount_input.replace(',', '')) if payment_amount_input else 2410
        
        duration_input = input("Enter Loan Duration in years [default: 30]: ").strip()
        duration_years = int(duration_input) if duration_input else 30
        
        interest_input = input("Enter Annual Interest Rate [default: 6.24%]: ").strip()
        annual_interest_rate = (float(interest_input.replace('%', '')) if interest_input else 6.24) / 100
        
        offset_start_input = input("Enter Offset Starting Balance [default: $20,000]: $").strip()
        offset_start_balance = float(offset_start_input.replace(',', '')) if offset_start_input else 20000
        
        offset_growth_input = input("Enter Offset Growth Per Month [default: $3,000]: $").strip()
        offset_growth_per_month = float(offset_growth_input.replace(',', '')) if offset_growth_input else 3000
        
        loan_start_input = input("Enter Loan Start Date (DD/MM/YYYY) [default: 29/08/2025]: ").strip()
        loan_start_date_str = loan_start_input if loan_start_input else "29/08/2025"

    except ValueError:
        print("Invalid input. Please ensure all numerical inputs are valid numbers.")
        exit()

    # Calculate and generate the schedule
    try:
        schedule_df = calculate_mortgage_schedule(
            loan_amount,
            payment_type,
            payment_amount,
            duration_years,
            annual_interest_rate,
            offset_start_balance,
            offset_growth_per_month,
            loan_start_date_str
        )

        # Save to Excel with proper formatting - include timestamp in filename
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"mortgage_schedule_{timestamp}.xlsx"
        
        # Create Excel writer object for custom formatting
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            # Calculate summary information using unformatted data
            unformatted_df = schedule_df.unformatted
            total_amount_paid = unformatted_df[f"{payment_type} Payment"].sum()
            payoff_time = unformatted_df['Remaining Term (Years & Months)'].iloc[0]
            
            # Write the schedule starting from row 18 to leave space for separator
            schedule_df.to_excel(writer, sheet_name='Mortgage Schedule', index=False, startrow=17)
            
            # Get the workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Mortgage Schedule']
            
            # Add summary information at the top
            from openpyxl.styles import Font, Alignment
            
            # Title
            worksheet['A1'] = 'MORTGAGE AMORTIZATION SCHEDULE WITH OFFSET ACCOUNT'
            worksheet['A1'].font = Font(bold=True, size=14)
            worksheet['A1'].alignment = Alignment(horizontal='center')
            worksheet.merge_cells('A1:G1')
            
            # Parameters (in black)
            row = 3
            worksheet[f'A{row}'] = 'LOAN PARAMETERS:'
            worksheet[f'A{row}'].font = Font(bold=True)
            
            row += 1
            worksheet[f'A{row}'] = f'Loan Amount:'
            worksheet[f'B{row}'] = f'${loan_amount:,.2f}'
            
            row += 1
            worksheet[f'A{row}'] = f'Payment Type:'
            worksheet[f'B{row}'] = payment_type
            
            row += 1
            worksheet[f'A{row}'] = f'{payment_type} Payment:'
            worksheet[f'B{row}'] = f'${payment_amount:,.2f}'
            
            row += 1
            worksheet[f'A{row}'] = f'Loan Duration:'
            worksheet[f'B{row}'] = f'{duration_years} years'
            
            row += 1
            worksheet[f'A{row}'] = f'Annual Interest Rate:'
            worksheet[f'B{row}'] = f'{annual_interest_rate*100:.2f}%'
            
            row += 1
            worksheet[f'A{row}'] = f'Offset Starting Balance:'
            worksheet[f'B{row}'] = f'${offset_start_balance:,.2f}'
            
            row += 1
            worksheet[f'A{row}'] = f'Offset Growth Per Month:'
            worksheet[f'B{row}'] = f'${offset_growth_per_month:,.2f}'
            
            row += 1
            worksheet[f'A{row}'] = f'Loan Start Date:'
            worksheet[f'B{row}'] = loan_start_date_str
            
            # Results (in red) - positioned separately from the table
            row += 2
            worksheet[f'A{row}'] = 'RESULTS:'
            worksheet[f'A{row}'].font = Font(bold=True, color='FF0000', size=12)
            
            row += 1
            worksheet[f'A{row}'] = f'Loan Paid Off In:'
            worksheet[f'A{row}'].font = Font(color='FF0000', bold=True)
            worksheet[f'B{row}'] = payoff_time
            worksheet[f'B{row}'].font = Font(color='FF0000', bold=True, size=11)
            
            row += 1
            worksheet[f'A{row}'] = f'Total Amount of Repayments:'
            worksheet[f'A{row}'].font = Font(color='FF0000', bold=True)
            worksheet[f'B{row}'] = f'${total_amount_paid:,.2f}'
            worksheet[f'B{row}'].font = Font(color='FF0000', bold=True, size=11)
            
            # Add a clear separator line before the table
            row += 2
            worksheet[f'A{row}'] = '=' * 80
            worksheet[f'A{row}'].font = Font(bold=True)
            worksheet.merge_cells(f'A{row}:H{row}')
            
            # Add one more blank row for spacing
            row += 1
            
            # Auto-adjust column widths (skip merged cells)
            for column in worksheet.columns:
                max_length = 0
                column_letter = None
                for cell in column:
                    try:
                        # Skip merged cells
                        if hasattr(cell, 'column_letter'):
                            column_letter = cell.column_letter
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                    except:
                        pass
                if column_letter and max_length > 0:
                    adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
                    worksheet.column_dimensions[column_letter].width = adjusted_width

        print(f"\nMortgage schedule generated successfully! Saved to '{output_filename}'")
        print(f"Loan paid off in: {schedule_df['Remaining Term (Years & Months)'].iloc[0]}")
        
    except Exception as e:
        print(f"\nAn error occurred: {e}")

# ```