#pragma once

#include <iostream>
#include <stdlib.h>
#include <msclr/marshal_cppstd.h>
#include <Windows.h>

// windows script encoding namespace
namespace windowsscriptencoding 
{
	// namespaces
	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	// class definition
	public ref class mainform : public System::Windows::Forms::Form
	{
		// public functions
		public:
			// run main form
			mainform(void)
			{
				// initialize components
				InitializeComponent();
			}

		// protected functions
		protected:
			// mainform proteted
			~mainform()
			{
				// check if components are null
				if (components)
					// delete if yes
					delete components;
			}

		// all components
		private: System::Windows::Forms::Label^ typelabel;
		private: System::Windows::Forms::Button^ vbsbutton;
		private: System::Windows::Forms::Button^ jsbutton;
		private: System::Windows::Forms::Label^ pathlabel;
		private: System::Windows::Forms::RichTextBox^ pathbox;
		private: System::Windows::Forms::Button^ browsebutton;
		private: System::Windows::Forms::Button^ encodebutton;
		private: System::Windows::Forms::Label^ outputdirlabel;
		private: System::Windows::Forms::RichTextBox^ outputdirbox;
		private: System::Windows::Forms::Button^ browsedirbutton;
		private: System::Windows::Forms::Button^ aboutbutton;

		// components container
		private:
			System::ComponentModel::Container ^components;

	// designer region
	#pragma region Windows Form Designer generated code

		// initialize
		void InitializeComponent(void)
		{
			// component manager
			System::ComponentModel::ComponentResourceManager^ resources = (gcnew System::ComponentModel::ComponentResourceManager(mainform::typeid));
			this->typelabel = (gcnew System::Windows::Forms::Label());
			this->vbsbutton = (gcnew System::Windows::Forms::Button());
			this->jsbutton = (gcnew System::Windows::Forms::Button());
			this->pathlabel = (gcnew System::Windows::Forms::Label());
			this->pathbox = (gcnew System::Windows::Forms::RichTextBox());
			this->browsebutton = (gcnew System::Windows::Forms::Button());
			this->encodebutton = (gcnew System::Windows::Forms::Button());
			this->outputdirlabel = (gcnew System::Windows::Forms::Label());
			this->outputdirbox = (gcnew System::Windows::Forms::RichTextBox());
			this->browsedirbutton = (gcnew System::Windows::Forms::Button());
			this->aboutbutton = (gcnew System::Windows::Forms::Button());
			this->SuspendLayout();
			
			// typelabel
			this->typelabel->AutoSize = true;
			this->typelabel->Location = System::Drawing::Point(9, 14);
			this->typelabel->Name = L"typelabel";
			this->typelabel->Size = System::Drawing::Size(99, 13);
			this->typelabel->TabIndex = 0;
			this->typelabel->Text = L"Scripting Language";

			// vbsbutton
			this->vbsbutton->Enabled = false;
			this->vbsbutton->Location = System::Drawing::Point(12, 30);
			this->vbsbutton->Name = L"vbsbutton";
			this->vbsbutton->Size = System::Drawing::Size(187, 51);
			this->vbsbutton->TabIndex = 1;
			this->vbsbutton->Text = L"VBScript";
			this->vbsbutton->UseVisualStyleBackColor = true;
			this->vbsbutton->Click += gcnew System::EventHandler(this, &mainform::vbsbutton_Click);

			// jsbutton
			this->jsbutton->Location = System::Drawing::Point(205, 30);
			this->jsbutton->Name = L"jsbutton";
			this->jsbutton->Size = System::Drawing::Size(171, 51);
			this->jsbutton->TabIndex = 2;
			this->jsbutton->Text = L"JScript";
			this->jsbutton->UseVisualStyleBackColor = true;
			this->jsbutton->Click += gcnew System::EventHandler(this, &mainform::jsbutton_Click);

			// pathlabel
			this->pathlabel->AutoSize = true;
			this->pathlabel->Location = System::Drawing::Point(9, 96);
			this->pathlabel->Name = L"pathlabel";
			this->pathlabel->Size = System::Drawing::Size(60, 13);
			this->pathlabel->TabIndex = 3;
			this->pathlabel->Text = L"Path to File";

			// pathbox
			this->pathbox->Location = System::Drawing::Point(12, 112);
			this->pathbox->Name = L"pathbox";
			this->pathbox->Size = System::Drawing::Size(283, 21);
			this->pathbox->TabIndex = 4;
			this->pathbox->Text = L"";
			this->pathbox->TextChanged += gcnew System::EventHandler(this, &mainform::textbox_TextChanged);

			// browsebutton
			this->browsebutton->Location = System::Drawing::Point(301, 112);
			this->browsebutton->Name = L"browsebutton";
			this->browsebutton->Size = System::Drawing::Size(75, 23);
			this->browsebutton->TabIndex = 5;
			this->browsebutton->Text = L"Browse";
			this->browsebutton->UseVisualStyleBackColor = true;
			this->browsebutton->Click += gcnew System::EventHandler(this, &mainform::browsebutton_Click);

			// encodebutton
			this->encodebutton->Enabled = false;
			this->encodebutton->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 9.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->encodebutton->Location = System::Drawing::Point(394, 112);
			this->encodebutton->Name = L"encodebutton";
			this->encodebutton->Size = System::Drawing::Size(296, 74);
			this->encodebutton->TabIndex = 9;
			this->encodebutton->Text = L"Encode Script";
			this->encodebutton->UseVisualStyleBackColor = true;
			this->encodebutton->Click += gcnew System::EventHandler(this, &mainform::encodebutton_Click);

			// outputdirlabel
			this->outputdirlabel->AutoSize = true;
			this->outputdirlabel->Location = System::Drawing::Point(12, 149);
			this->outputdirlabel->Name = L"outputdirlabel";
			this->outputdirlabel->Size = System::Drawing::Size(84, 13);
			this->outputdirlabel->TabIndex = 10;
			this->outputdirlabel->Text = L"Output Directory";

			// outputdirbox
			this->outputdirbox->Location = System::Drawing::Point(12, 165);
			this->outputdirbox->Name = L"outputdirbox";
			this->outputdirbox->Size = System::Drawing::Size(283, 21);
			this->outputdirbox->TabIndex = 11;
			this->outputdirbox->Text = L"";
			this->outputdirbox->TextChanged += gcnew System::EventHandler(this, &mainform::textbox_TextChanged);

			// browsedirbutton
			this->browsedirbutton->Location = System::Drawing::Point(301, 163);
			this->browsedirbutton->Name = L"browsedirbutton";
			this->browsedirbutton->Size = System::Drawing::Size(74, 23);
			this->browsedirbutton->TabIndex = 12;
			this->browsedirbutton->Text = L"Browse";
			this->browsedirbutton->UseVisualStyleBackColor = true;
			this->browsedirbutton->Click += gcnew System::EventHandler(this, &mainform::browsedirbutton_Click);

			// aboutbutton
			this->aboutbutton->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 9.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->aboutbutton->Location = System::Drawing::Point(394, 30);
			this->aboutbutton->Name = L"aboutbutton";
			this->aboutbutton->Size = System::Drawing::Size(296, 69);
			this->aboutbutton->TabIndex = 13;
			this->aboutbutton->Text = L"About";
			this->aboutbutton->UseVisualStyleBackColor = true;
			this->aboutbutton->Click += gcnew System::EventHandler(this, &mainform::aboutbutton_Click);

			// mainform
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(702, 197);
			this->Controls->Add(this->aboutbutton);
			this->Controls->Add(this->browsedirbutton);
			this->Controls->Add(this->outputdirbox);
			this->Controls->Add(this->outputdirlabel);
			this->Controls->Add(this->encodebutton);
			this->Controls->Add(this->browsebutton);
			this->Controls->Add(this->pathbox);
			this->Controls->Add(this->pathlabel);
			this->Controls->Add(this->jsbutton);
			this->Controls->Add(this->vbsbutton);
			this->Controls->Add(this->typelabel);
			this->FormBorderStyle = System::Windows::Forms::FormBorderStyle::FixedSingle;
			this->Icon = (cli::safe_cast<System::Drawing::Icon^>(resources->GetObject(L"$this.Icon")));
			this->Name = L"mainform";
			this->Text = L"Windows Script Encoding";
			this->ResumeLayout(false);
			this->PerformLayout();
			this->Load += gcnew System::EventHandler(this, &mainform::document_Load);

		}

	// end that region
	#pragma endregion

	// functions region
	#pragma region Functions

		// on load
		private: System::Void document_Load(System::Object^ sender, System::EventArgs^ e)
		{
			// get message box variables
			String^ message = "This program is in development.\nIf you see any issues in this program, please report them under the \"issues\" section of the github page.";
			String^ title = "Notice";
			MessageBoxButtons buttons = MessageBoxButtons::OK;
			MessageBoxIcon icon = MessageBoxIcon::Information;

			// show
			System::Windows::Forms::DialogResult result = System::Windows::Forms::MessageBox::Show(message, title, buttons, icon);
		}

		// button click
		private: System::Void aboutbutton_Click(System::Object^ sender, System::EventArgs^ e)
		{
			// get message box variables
			String^ message = "This program was developed by Yochran.\nWould you like to open the GitHub page for this program?";
			String^ title = "About this Program";
			MessageBoxButtons buttons = MessageBoxButtons::YesNo;
			MessageBoxIcon icon = MessageBoxIcon::Information;

			// show
			System::Windows::Forms::DialogResult result = System::Windows::Forms::MessageBox::Show(message, title, buttons, icon);

			// check clicked button
			if (result == System::Windows::Forms::DialogResult::Yes)
				// open browser
				system("start https://github.com/Yochran/windows-script-encoding");
		}

		// button click
		private: System::Void vbsbutton_Click(System::Object^ sender, System::EventArgs^ e)
		{
			// switch selection
			this->jsbutton->Enabled = true;
			this->vbsbutton->Enabled = false;
		}

		// button click
		private: System::Void jsbutton_Click(System::Object^ sender, System::EventArgs^ e)
		{
			// switch selection
			this->jsbutton->Enabled = false;
			this->vbsbutton->Enabled = true;
		}

		// textbox change
		private: System::Void textbox_TextChanged(System::Object^ sender, System::EventArgs^ e) 
		{
			// check if textbox is set
			if (this->pathbox->Text == "" || this->pathbox->Text == " " || this->outputdirbox->Text == "" || this->outputdirbox->Text == " ")
				// disable
				this->encodebutton->Enabled = false;
			else
				// enable
				this->encodebutton->Enabled = true;
		}

		// button click
		private: System::Void browsebutton_Click(System::Object^ sender, System::EventArgs^ e)
		{
			// open folder browser
			OpenFileDialog browser;
			browser.InitialDirectory = "C:\\";
			browser.Filter = "VBScript files (*.vbs)|*.vbs|JScript files (*.js)|*.js|All files (*.*)|*.*";
			browser.FilterIndex = 3;
			browser.RestoreDirectory = true;

			// check if button clicked was "open"
			if (browser.ShowDialog() == System::Windows::Forms::DialogResult::OK)
				this->pathbox->Text = browser.FileName;
		}

		// button click
		private: System::Void browsedirbutton_Click(System::Object^ sender, System::EventArgs^ e) 
		{
			FolderBrowserDialog browser;
			browser.Description = "Select a Folder";

			// check if button clicked was "ok"
			if (browser.ShowDialog() == System::Windows::Forms::DialogResult::OK)
				this->outputdirbox->Text = browser.SelectedPath;
		}

		// button click
		private: System::Void encodebutton_Click(System::Object^ sender, System::EventArgs^ e)
		{
			// get path to encoder script
			std::string path = "D:\\vsc\\windows-script-encoding\\encoder.vbs";

			// create context
			msclr::interop::marshal_context context;

			// get file
			std::string file = context.marshal_as<std::string>(this->pathbox->Text);

			// create extension & type variables
			std::string in_extension, out_extension, type;

			// set the type
			if (this->jsbutton->Enabled == true)
			{
				type = "VBScript";
				in_extension = ".vbs";
				out_extension = ".vbe";
			}
			else if (this->vbsbutton->Enabled == true)
			{
				type = "JScript";
				in_extension = ".js";
				out_extension = ".jse";
			}

			// get the output directory & file
			std::string output_dir = context.marshal_as<std::string>(this->outputdirbox->Text);
			std::string output = output_dir + "\\encoded" + out_extension;

			// create command
			std::string command = "cscript encoder.vbs " + file + " " + in_extension + " " + type + " " + output;

			// run encoder script
			system(command.c_str());
		}
	};

	// end that region
	#pragma endregion
}
