# Meinhardt CV Automation

## Overview

**Meinhardt CV Automation** is a web-based application designed to streamline and automate the process of converting candidate resumes into standardized CV documents that match a client's specific format and template requirements.

This tool eliminates manual data entry and formatting by intelligently extracting information from candidate resumes and automatically populating a pre-designed Word document template with extracted data. It leverages AI technology to ensure accuracy and consistency across all generated CVs.

---

## Purpose & Use Case

### The Problem
- Recruiters and HR teams often need to convert multiple candidate resumes into a standardized format required by clients
- Manual conversion is time-consuming and prone to errors
- Different candidates provide information in different formats and structures
- Maintaining consistent formatting across multiple CVs is challenging

### The Solution
- Upload a client's CV template (Word document)
- Upload multiple candidate resumes (PDF or Word files)
- The system automatically extracts key information and populates the template
- Download standardized CVs ready for delivery to the client

---

## Key Features

### 1. **Template-Based Generation**
- Supports custom client CV formats
- Uses Word documents as templates
- Preserves client branding, layout, and structure
- One template can be used for all candidates

### 2. **Multi-Format Resume Support**
- Accepts PDF files
- Accepts Word documents (.docx)
- Handles various resume layouts and formats

### 3. **AI-Powered Data Extraction**
- Uses Google Gemini API for intelligent information extraction
- Identifies key CV fields (name, contact, experience, education, skills, etc.)
- Includes confidence scoring for extracted data
- Falls back to rule-based extraction if AI is unavailable

### 4. **Batch Processing**
- Process one resume or multiple resumes simultaneously
- Generate individual Word files or download as ZIP archive
- Saves time for large recruitment campaigns

### 5. **Data Accuracy**
- Configurable confidence thresholds
- Uncertain data is flagged for manual review
- Extraction results are saved as JSON for audit trail

### 6. **Integration Capabilities**
- Local REST API endpoint for workflow automation
- Compatible with n8n automation platform
- Can be triggered from external tools and systems

---

## How It Works

### Workflow

```
1. USER UPLOADS TEMPLATE
   ↓
2. USER UPLOADS RESUMES (PDF/DOCX)
   ↓
3. SYSTEM EXTRACTS DATA FROM EACH RESUME
   - Uses Gemini AI or falls back to local rules
   ↓
4. SYSTEM POPULATES TEMPLATE
   - Inserts extracted data into template fields
   ↓
5. USER DOWNLOADS GENERATED CVS
   - Individual files or ZIP archive
   ↓
6. OUTPUT SAVED
   - Generated files in outputs folder
   - Extraction data in JSON format
```

### Data Flow
- **Uploads**: Stored in `uploads/` directory
- **Generated CVs**: Stored in `outputs/` directory with unique timestamps
- **Extraction Results**: JSON files capture all extracted data and confidence scores
- **Source Files**: Preserved for audit and troubleshooting

---

## Key Components

### Web Interface
- User-friendly dashboard
- File upload area
- Generation status tracking
- Download options

### Processing Engine
- Resume parsing and data extraction
- Template population logic
- Format preservation
- Error handling and fallback mechanisms

### AI Integration
- Google Gemini API for primary extraction
- Local rule-based fallback extraction
- Confidence scoring system
- Structured data output

### Storage & Outputs
- Uploaded files management
- Generated document storage
- Extraction metadata and audit logs
- ZIP archive creation for bulk downloads

---

## System Architecture

### Technology Stack
- **Backend**: Python
- **Frontend**: Web-based interface
- **Document Processing**: python-docx library
- **PDF Handling**: pypdf library
- **AI Provider**: Google Gemini API
- **Server**: Local HTTP server on port 8765

### Configuration
- Environment-based settings for API keys
- Configurable confidence thresholds
- Customizable timeout settings
- Support for multiple Gemini models

---

## Getting Started

### Prerequisites
- Python environment
- Google Gemini API key
- Modern web browser
- Client CV template in Word format

### Quick Start
1. Install dependencies from `requirements.txt`
2. Configure Google Gemini API key (see setup documentation)
3. Run the application server
4. Open web interface in browser
5. Upload template and resumes
6. Generate CVs
7. Download results

### Setup Resources
- **Detailed Setup**: See `docs/MANUAL_STEPS.md`
- **Environment Configuration**: See `env.md`
- **API Setup Guide**: Instructions for Google Gemini API key creation

---

## Automation & Integration

### n8n Workflow Integration
- Pre-configured n8n workflow template included
- Enables automated CV generation from external triggers
- REST API endpoint: `http://127.0.0.1:8765/api/process`
- Supports manual trigger workflows and automated pipelines

### Use Cases for Automation
- Scheduled batch processing of new candidates
- Integration with Applicant Tracking Systems (ATS)
- Email-triggered CV generation
- Integration with recruitment databases
- Multi-step recruitment workflows

---

## Data & Output Management

### Generated Files
- Individual Word documents per candidate
- ZIP archives containing all generated CVs
- Extraction result JSON files for review
- Timestamped output folders for organization

### Stored Data
- Uploaded templates and resumes
- Extraction metadata and confidence scores
- Processing logs and error messages
- Audit trail of all operations

### Data Quality
- Missing or uncertain fields flagged in JSON output
- Confidence scores for each extracted field
- Manual review capability for unconfident extractions
- Validation against field requirements

---

## Accuracy & Quality Assurance

### Confidence Scoring
- Each extracted field receives a confidence score
- Fields below threshold are left blank
- Low-confidence extractions documented in JSON
- Enables quality control before delivery

### Fallback Mechanisms
- AI-first approach with local extraction fallback
- Graceful degradation if API unavailable
- Consistent output even with service interruptions
- Error reporting for audit purposes

### Review & Validation
- Generated CVs can be manually reviewed before sending
- Extraction JSON available for detailed inspection
- Field-by-field confidence transparency
- Easy correction workflow

---

## Security & Privacy

### Key Management
- Gemini API key never logged or exposed
- Environment variable-based configuration
- Support for secure key storage
- Instructions for key rotation and management

### Data Handling
- Uploaded files stored locally
- No external data transmission except to Gemini API
- Option to delete processed files after completion
- Audit logs for compliance tracking

### Best Practices
- Keep API keys secure and private
- Rotate keys if exposure suspected
- Review extraction results before delivery
- Maintain audit logs for regulatory compliance

---

## Troubleshooting & Support

### Common Issues
- **Missing API Key**: Application will use local fallback extraction
- **Invalid Template Format**: Template must be .docx format
- **Resume Format Issues**: Ensure PDFs are text-based, not scanned images
- **Timeout Errors**: Adjust timeout settings for large files

### Getting Help
- Review `docs/MANUAL_STEPS.md` for step-by-step guidance
- Check `env.md` for environment setup issues
- Examine extraction JSON files for data-related problems
- Verify API quota and rate limits with Gemini

### Monitoring
- Check output folder for generated files
- Review extraction JSON for quality metrics
- Monitor API usage and quota
- Track processing times for performance optimization

---

## Maintenance & Updates

### Regular Tasks
- Clean up old output files periodically
- Monitor API quota and usage
- Review extraction accuracy metrics
- Update dependencies as needed

### Performance Optimization
- Monitor processing times
- Adjust confidence thresholds based on results
- Review and optimize template structure
- Cache extraction results when possible

### Backup & Recovery
- Maintain backups of CV templates
- Archive important generated files
- Keep extraction metadata for audits
- Document template customizations

---

## Next Steps

1. **Setup**: Follow the setup guide in `docs/MANUAL_STEPS.md`
2. **Configuration**: Set up your Gemini API key using `env.md`
3. **First Run**: Test with provided sample files
4. **Template Customization**: Adapt the CV template for your clients
5. **Integration**: Set up n8n workflow if using automation
6. **Deployment**: Deploy to your preferred infrastructure

---

## Additional Resources

- **Setup Instructions**: `docs/MANUAL_STEPS.md`
- **Environment Configuration**: `env.md`
- **API Provider Details**: `solution.md`
- **n8n Workflow Template**: `docs/n8n_workflow.json`

---

## Support & Feedback

For technical issues, refer to the documentation files included in the project. For feature requests or improvements, review the current implementation and adapt as needed for your specific use case.

---

**Last Updated**: May 2026  
**Version**: 1.0 - Gemini API Integration
