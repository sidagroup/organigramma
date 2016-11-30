<?php

class SiteController extends Controller
{
	/**
	 * Declares class-based actions.
	 */
	public function actions()
	{
		return array(
			// captcha action renders the CAPTCHA image displayed on the contact page
			'captcha'=>array(
				'class'=>'CCaptchaAction',
				'backColor'=>0xFFFFFF,
			),
			// page action renders "static" pages stored under 'protected/views/site/pages'
			// They can be accessed via: index.php?r=site/page&view=FileName
			'page'=>array(
				'class'=>'CViewAction',
			),
		);
	}

	/**
	 * This is the default 'index' action that is invoked
	 * when an action is not explicitly requested by users.
	 */
	public function actionIndex()
	{

		 //ARRAY ASSOCIATIVO CHE CONTIENE LA LISTA DEI CORDINATORI, DEI PMA E DEI PMS
		 $objData = array();
		
		 $persone = Persons::model(); 
		

		 //QUERY MANAGER
		 $criteriaM = new CDbCriteria();
		 $criteriaM->condition='RoleID=5';		 
		 $manager = $persone->find($criteriaM);

		 //QUERY COORDINATORI
		 $criteriaC = new CDbCriteria();
		 $criteriaC->with = array('personsMasters.master');
		 $criteriaC->together=true;
		 $criteriaC->condition='RoleID=18 AND t.Enabled=1 AND master.Enabled=1';		 
		 $coordinatori = $persone->findAll($criteriaC);

         //ciclo la lista dei cordinatori
		 foreach ($coordinatori as $coordinatore) {
		 	
		 	//lista dei master associata al coordinatore
        	$arrayMastersCoordinatore = array();
        	foreach ($coordinatore->personsMasters as $master) {
        		$arrayMastersCoordinatore[]=$master->MasterID;        		
        	}

        	 //QUERY MARKETING ASSOCIATI AL COORDINATORE
	 		 $criteriaMktg = new CDbCriteria();
	 		 $criteriaMktg->with = array('personsMasters.master');
			 $criteriaMktg->together=true;
			 $criteriaMktg->addCondition('RoleID=19 AND t.Enabled=1');
			 $criteriaMktg->addInCondition('personsMasters.MasterID',$arrayMastersCoordinatore);
			 $listamktg = $persone->findAll($criteriaMktg);

			 //QUERY DIDATTICA ASSOCIATI AL COORDINATORE
	 		 $criteriaDid = new CDbCriteria();
	 		 $criteriaDid->with = array('personsMasters.master');
			 $criteriaDid->together=true;
			 $criteriaDid->addCondition('RoleID=7 AND t.Enabled=1');
			 $criteriaDid->addInCondition('personsMasters.MasterID',$arrayMastersCoordinatore);
			 $listadid = $persone->findAll($criteriaDid);

			 //QUERY RPA ASSOCIATI AL COORDINATORE
	 		 $criteriaRpa = new CDbCriteria();
	 		 $criteriaRpa->with = array('personsMasters.master');
			 $criteriaRpa->together=true;
			 $criteriaRpa->addCondition('RoleID=13 AND t.Enabled=1');
			 $criteriaRpa->addInCondition('personsMasters.MasterID',$arrayMastersCoordinatore);
			 $listarpa = $persone->findAll($criteriaRpa);

        	 //QUERY PMA ASSOCIATI AL COORDINATORE
	 		 $criteriaPma = new CDbCriteria();
	 		 $criteriaPma->with = array('personsMasters.master');
			 $criteriaPma->together=true;
			 $criteriaPma->addCondition('RoleID=11 AND t.Enabled=1');
			 $criteriaPma->addInCondition('personsMasters.MasterID',$arrayMastersCoordinatore);
			 $listapma = $persone->findAll($criteriaPma);
			
			 //QUERY PMS DI ANCONA ASSOCIATI AL COORDINATORE
			 $criteriaPms = new CDbCriteria();
	 		 $criteriaPms->with = array('personsMasters.master','personsCities.city');
			 $criteriaPms->together=true;
			 $criteriaPms->addCondition('RoleID=15 AND t.Enabled=1 AND personsCities.CityID=10');
			 $criteriaPms->addInCondition('personsMasters.MasterID',$arrayMastersCoordinatore);

			 $listapms = $persone->findAll($criteriaPms);

			 //QUERY PMS CHE NON SONO DI ANCONA  ASSOCIATI AL COORDINATORE
			 $criteriaPmsOutsider = new CDbCriteria();
	 		 $criteriaPmsOutsider->with = array('personsMasters.master','personsCities.city');
			 $criteriaPmsOutsider->together=true;
			 $criteriaPmsOutsider->addCondition('RoleID=15 AND t.Enabled=1 AND personsCities.CityID!=10 AND city.Enabled=1');
			 $criteriaPmsOutsider->addInCondition('personsMasters.MasterID',$arrayMastersCoordinatore);
			 $criteriaPmsOutsider->order = 'city.Name';
			 $listapmsOutsider = $persone->findAll($criteriaPmsOutsider);

			 
			//COSTRUISCO l' OBJECTDATA
			$objData[$coordinatore->PersonID] = array($coordinatore,$listapma,$listapms,$listapmsOutsider,$listamktg,$listadid,$listapma);
        	
		 }

			

		
		//DISEGNA L'ORGANIGRAMMA A PARTIRE DA OBJDATA
		$this->draw($objData,$manager);
	
	}


	public function draw($objData,$manager){

		// INIT
		$posXCoordinatore = 0;
		$posYCoordinatore = 13;
		//SPAZIATURA VERTICALE TRA COORDINATORI PMA/PMS ANCONA E PMS OUTSIDER
		$spazioY = 2;

		//DIMENSIONE DELLE CELLE DI DEFAULT
		$cellWidth = 5;

		//DESCRIZION DEI MASTER (NORMAL O SHORT)
		//$descrizione = 'Description';
		$descrizione = 'ShortDescription';


		//STYLE DELLE CELLE
		$styleManager = array(	
				'fill'=>array(
					'type'=>PHPExcel_Style_Fill::FILL_SOLID,
					'startcolor'=>array(
						'argb'=>'FF6666'),
					),
				);
		$styleCoordinatore = array(	
				'fill'=>array(
					'type'=>PHPExcel_Style_Fill::FILL_SOLID,
					'startcolor'=>array(
						'argb'=>'FFFF00'),
					),
				);

		$stylePma = array(
				'fill'=>array(
					'type'=>PHPExcel_Style_Fill::FILL_SOLID,
					'startcolor'=>array(
						'argb'=>'E0FFFF'),
					),
				);

		$styleAree = array(
				'fill'=>array(
					'type'=>PHPExcel_Style_Fill::FILL_SOLID,
					'startcolor'=>array(
						'argb'=>'FFE5CC'),
					),
				);

		$stylePms = array(
				'fill'=>array(
					'type'=>PHPExcel_Style_Fill::FILL_SOLID,
					'startcolor'=>array(
						'argb'=>'E0FFBF'),
					),
				);

		$styleTitle = array(
				'borders'=>array(
						'top'=>array(
							'style'=>PHPExcel_Style_Border::BORDER_THIN
							),
						'left'=>array(
							'style'=>PHPExcel_Style_Border::BORDER_THIN
							),
						'right'=>array(
							'style'=>PHPExcel_Style_Border::BORDER_THIN
							),
						),
				'font'=>array(
					'bold'=>true,
					),
				'alignment' => array(
       				'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
       				'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
       				),
				);

		$styleCorsi = array(
				'borders'=>array(
						'outline'=>array(
							'style'=>PHPExcel_Style_Border::BORDER_THIN
							),
						),
				'font'=>array(
					'italic'=>true,
					'size'=>6,
					),
				'alignment' => array(
	       			'wrap' => true,
	       			'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
	       			'vertical' => PHPExcel_Style_Alignment::VERTICAL_BOTTOM,
	       			),
				);

		$styleNomi = array(
			'borders'=>array(
						'bottom'=>array(
							'style'=>PHPExcel_Style_Border::BORDER_THIN
							),
						'left'=>array(
							'style'=>PHPExcel_Style_Border::BORDER_THIN
							),
						'right'=>array(
							'style'=>PHPExcel_Style_Border::BORDER_THIN
							),
						),
			'alignment' => array(
       			'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
       			'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
       			),
			);


		$objPHPExcel = new PHPExcel();
       
        $sheet = $objPHPExcel->getActiveSheet()->setTitle('Organigramma SIDA');

  		//CICLO TRA OBJDATA 
  		//OGNI CHIAVE è L'ID DEL COORDINATORE
  		//VALUE[0] = COORDINATORE
  		//VALUE[1] = LISTA DEI PMA
  		//VALUE[2] = LISTA DEI PMS ANCONA
  		//VALUE[3] = LISTA DEI PMS FUORI ANCONA
  		//VALUE[4] = LISTA DEI PMS FUORI ANCONA
  		//VALUE[5] = LISTA DEI PMS FUORI ANCONA
  		//VALUE[6] = LISTA DEI PMS FUORI ANCONA
  		$posXLastCoord = 0;
        $acc = 0;
        foreach ($objData as $key => $value) {
       	    
       		$numeroPma = count($value[1]);
       		$posXCoordinatore = $posXCoordinatore + floor($numeroPma/2) + 1 + $acc;
       		$posXLastCoord = $posXCoordinatore;
       		$acc = floor($numeroPma/2) + 1 ;

       		$sheet->setCellValueByColumnAndRow($posXCoordinatore, $posYCoordinatore,  "COORDINATORE");
       		$sheet->setCellValueByColumnAndRow($posXCoordinatore, $posYCoordinatore+1,  ucfirst(strtolower($value[0]->FirstName)) . " " . ucfirst(strtolower($value[0]->LastName)) );

       		//SETTO GLI STILI ALLE CELLE
       		$el =  $sheet->getColumnDimensionByColumn($posXCoordinatore);
       		$el->setAutoSIze(true);

       		$el = $sheet->getStyleByColumnAndRow($posXCoordinatore, $posYCoordinatore);
       		$el->applyFromArray($styleCoordinatore);
        	$el->applyFromArray($styleTitle);

        	$el = $sheet->getStyleByColumnAndRow($posXCoordinatore, $posYCoordinatore+1);
       		$el->applyFromArray($styleCoordinatore);
       		$el->applyFromArray($styleNomi);
        

       		
       		//STAMPIAMO l'AREA MARKETING
       		$posXmktg = $posXCoordinatore -1;
       		$posYmktg = $posYCoordinatore + 1 + $spazioY;

       		foreach ($value[4] as $mktg) {
				
       			$sheet->setCellValueByColumnAndRow( $posXmktg , $posYmktg ,  "AREA MARKETING");
       			$sheet->setCellValueByColumnAndRow($posXmktg , $posYmktg + 1, ucfirst(strtolower($mktg->FirstName)) . " " . ucfirst(strtolower($mktg->LastName)) );

       			//SETTO GLI STILI ALLE CELLE
       			$elpma =  $sheet->getColumnDimensionByColumn($posXmktg);
       			$elpma->setAutoSIze(true);

       			$elpma = $sheet->getStyleByColumnAndRow($posXmktg, $posYmktg);
       			$elpma->applyFromArray($styleAree);       			
        		$elpma->applyFromArray($styleTitle);

        		$elpma = $sheet->getStyleByColumnAndRow($posXmktg, $posYmktg + 1);
       			$elpma->applyFromArray($styleAree);
       			$elpma->applyFromArray($styleNomi);
       		}

       		//STAMPIAMO l'AREA DIDATTICA
       		$posXdid = $posXmktg;
       		$posYdid = $posYmktg + 1 + $spazioY;

       		foreach ($value[5] as $did) {
				
       			$sheet->setCellValueByColumnAndRow( $posXdid , $posYdid ,  "AREA DIDATTICA");
       			$sheet->setCellValueByColumnAndRow($posXdid , $posYdid + 1, ucfirst(strtolower($did->FirstName)) . " " . ucfirst(strtolower($did->LastName)) );

       			//SETTO GLI STILI ALLE CELLE
       			$elpma =  $sheet->getColumnDimensionByColumn($posXdid);
       			$elpma->setAutoSIze(true);

       			$elpma = $sheet->getStyleByColumnAndRow($posXdid, $posYdid);
       			$elpma->applyFromArray($styleAree);       			
        		$elpma->applyFromArray($styleTitle);

        		$elpma = $sheet->getStyleByColumnAndRow($posXdid, $posYdid + 1);
       			$elpma->applyFromArray($styleAree);
       			$elpma->applyFromArray($styleNomi);
       		}

       		//STAMPIAMO RPA
       		$posXrpa = $posXCoordinatore + 1;
       		$posYrpa = $posYCoordinatore + 1 + $spazioY;

       		foreach ($value[6] as $rpa) {
				
       			$sheet->setCellValueByColumnAndRow( $posXrpa , $posYrpa ,  "AREA PLACEMENT");
       			$sheet->setCellValueByColumnAndRow($posXrpa , $posYrpa + 1, ucfirst(strtolower($rpa->FirstName)) . " " . ucfirst(strtolower($rpa->LastName)) );

       			//SETTO GLI STILI ALLE CELLE
       			$elpma =  $sheet->getColumnDimensionByColumn($posXrpa);
       			$elpma->setAutoSIze(true);

       			$elpma = $sheet->getStyleByColumnAndRow($posXrpa, $posYrpa);
       			$elpma->applyFromArray($styleAree);       			
        		$elpma->applyFromArray($styleTitle);

        		$elpma = $sheet->getStyleByColumnAndRow($posXrpa, $posYrpa + 1);
       			$elpma->applyFromArray($styleAree);
       			$elpma->applyFromArray($styleNomi);
       		}

       		//STAMPIAMO I PMA
       		$posXpma = $posXCoordinatore - ($acc - 1);
       		$posYpma = $posYdid + 1 + $spazioY;

       		foreach ($value[1] as $pma) {
       			
       			//COSTRUISCO LA STRINGA DEI MASTER DEL PMA
       			$listaMasters = "";
       			foreach ($pma->personsMasters as $master) {
       				$listaMasters .= "\n".$master->master[$descrizione];
       			}

       			$sheet->setCellValueByColumnAndRow( $posXpma , $posYpma ,  "PMA");
       			$sheet->setCellValueByColumnAndRow($posXpma , $posYpma + 1, ucfirst(strtolower($pma->FirstName)) . " " . ucfirst(strtolower($pma->LastName)) );
       			$sheet->setCellValueByColumnAndRow($posXpma , $posYpma + 2, $listaMasters);
				
				//SETTO GLI STILI ALLE CELLE
       			$elpma =  $sheet->getColumnDimensionByColumn($posXpma);
       			$elpma->setAutoSIze(true);

       			$elpma = $sheet->getStyleByColumnAndRow($posXpma, $posYpma);
       			$elpma->applyFromArray($stylePma);       			
        		$elpma->applyFromArray($styleTitle);

        		$elpma = $sheet->getStyleByColumnAndRow($posXpma, $posYpma + 1);
       			$elpma->applyFromArray($stylePma);
       			$elpma->applyFromArray($styleNomi);
       			

        		$elpma = $sheet->getStyleByColumnAndRow($posXpma, $posYpma + 2);
       			$elpma->applyFromArray($stylePma);
        		$elpma->applyFromArray($styleCorsi);
       			
        		
        		//STAMPIAMO I PMS DI ANCONA
        		foreach ($pma->personsMasters as $masterpma) {
       				
       				foreach ($value[2] as $pmsAncona2) {

       					foreach ($pmsAncona2->personsMasters as $masterpms) {

       						//ASSOCIO UN  PMS A UN PMA TRAMITE I MASTER IN COUMUNE
       						if($masterpms->MasterID == $masterpma->MasterID){
       							
       							$sheet->setCellValueByColumnAndRow( $posXpma , $posYpma + 4 ,  "PMS ANCONA");
       							$sheet->setCellValueByColumnAndRow( $posXpma , $posYpma + 5, ucfirst(strtolower($pmsAncona2->FirstName)) . " " . ucfirst(strtolower($pmsAncona2->LastName)) );
	
								//SETTO GLI STILI ALLE CELLE
								$elpms = $sheet->getColumnDimensionByColumn($posXpma);
				       			$elpms->setAutoSIze(true);

				       			$elpms = $sheet->getStyleByColumnAndRow($posXpma , $posYpma + 4);
				       			$elpms->applyFromArray($stylePms);
				        		$elpms->applyFromArray($styleTitle);

				        		$elpms = $sheet->getStyleByColumnAndRow($posXpma , $posYpma + 5);
				       			$elpms->applyFromArray($stylePms);
				       			$elpms->applyFromArray($styleNomi);
						       			
       							break 2; //esco dal ciclo se trovo almeno un master uguale
       						}
       					} 
       					 
       				} 
       				
       			}

       			$posXpma +=1;
       		}

       		//STAMPIAMO I PMS CHE NON SONO DI ANCONA
       		$posXpmsOutsider = $posXCoordinatore;
       		$posYpmsOutsider = $posYpma + 6 + $spazioY;
       		       		
       		foreach ($value[3] as $pmsOutsider) {
       			$sheet->setCellValueByColumnAndRow( $posXpmsOutsider , $posYpmsOutsider ,  "PMS " . strtoupper($pmsOutsider->personsCities[0]->city->Name) );
       			$sheet->setCellValueByColumnAndRow($posXpmsOutsider , $posYpmsOutsider + 1, ucfirst(strtolower($pmsOutsider->FirstName)) . " " . ucfirst(strtolower($pmsOutsider->LastName)) );

       			//SETTO GLI STILI ALLE CELLE
       			$elpmsoutsider =  $sheet->getColumnDimensionByColumn($posXpmsOutsider);
       			$elpmsoutsider->setAutoSIze(true);

       			$elpmsoutsider = $sheet->getStyleByColumnAndRow($posXpmsOutsider, $posYpmsOutsider);
       			$elpmsoutsider->applyFromArray($stylePms);
        		$elpmsoutsider->applyFromArray($styleTitle);

        		$elpmsoutsider = $sheet->getStyleByColumnAndRow($posXpmsOutsider, $posYpmsOutsider + 1);
       			$elpmsoutsider->applyFromArray($stylePms);
       			$elpmsoutsider->applyFromArray($styleNomi);
       			
       			$posYpmsOutsider +=3;
       		}

       } //FINE DEL CICLO DI OBJDATA

       //STAMPO LA DATA DI CREAZIONE
       	date_default_timezone_set('Europe/Rome');
        $time = date("d/m/Y  H:i");
		$sheet->setCellValue('A8', 'Creato in data: '.  $time);
		$sheet->getStyle('A8')
		       ->getNumberFormat()
		       ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);
		 $sheet->getColumnDimensionByColumn('A')->setAutoSIze(true);
        
        //STAMPO IL LOGO DELLA SIDA
		$objDrawing = new PHPExcel_Worksheet_Drawing();
		$objDrawing->setName('Logo');
		$objDrawing->setDescription('Logo');
		$objDrawing->setPath(Yii::app()->basePath . '/../files/images/logo.png');
		$objDrawing->setCoordinates('A1');
		$objDrawing->setHeight(100);
		$objDrawing->setWorksheet($objPHPExcel->getActiveSheet());

		//MANAGER
		$posXmanager = $posXLastCoord/2+2;
		$sheet->setCellValueByColumnAndRow($posXmanager, 1,  "DIREZIONE");
		$sheet->setCellValueByColumnAndRow($posXLastCoord/2+2, 2,  $manager->FirstName . ' ' . $manager->LastName);
		//SETTO GLI STILI ALLE CELLE
   		$el =  $sheet->getColumnDimensionByColumn($posXmanager);
   		$el->setAutoSIze(true);

   		$el = $sheet->getStyleByColumnAndRow($posXmanager, 1);
   		$el->applyFromArray($styleManager);
    	$el->applyFromArray($styleTitle);

    	$el = $sheet->getStyleByColumnAndRow($posXmanager, 2);
   		$el->applyFromArray($styleManager);
   		$el->applyFromArray($styleNomi);

		//ALTRI STILI AL MIO FOGLIO EXCEL
		$sheet->setShowGridLines(false);
		$sheet->getSheetView()->setZoomScale(50);
		$sheet->getTabColor()->setRGB('007c53');

		//SCRITTURA SUL FILE
        $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
        $objWriter->save(Yii::app()->basePath . '/../files/exports/OrganigrammaSida.xlsx');

	}

	/**
	 * This is the action to handle external exceptions.
	 */
	public function actionError()
	{
		if($error=Yii::app()->errorHandler->error)
		{
			if(Yii::app()->request->isAjaxRequest)
				echo $error['message'];
			else
				$this->render('error', $error);
		}
	}



	/**
	 * Displays the contact page
	 */
	public function actionContact()
	{
		$model=new ContactForm;
		if(isset($_POST['ContactForm']))
		{
			$model->attributes=$_POST['ContactForm'];
			if($model->validate())
			{
				$name='=?UTF-8?B?'.base64_encode($model->name).'?=';
				$subject='=?UTF-8?B?'.base64_encode($model->subject).'?=';
				$headers="From: $name <{$model->email}>\r\n".
					"Reply-To: {$model->email}\r\n".
					"MIME-Version: 1.0\r\n".
					"Content-Type: text/plain; charset=UTF-8";

				mail(Yii::app()->params['adminEmail'],$subject,$model->body,$headers);
				Yii::app()->user->setFlash('contact','Thank you for contacting us. We will respond to you as soon as possible.');
				$this->refresh();
			}
		}
		$this->render('contact',array('model'=>$model));
	}

	/**
	 * Displays the login page
	 */
	public function actionLogin()
	{
		$model=new LoginForm;

		// if it is ajax validation request
		if(isset($_POST['ajax']) && $_POST['ajax']==='login-form')
		{
			echo CActiveForm::validate($model);
			Yii::app()->end();
		}

		// collect user input data
		if(isset($_POST['LoginForm']))
		{
			$model->attributes=$_POST['LoginForm'];
			// validate user input and redirect to the previous page if valid
			if($model->validate() && $model->login())
				$this->redirect(Yii::app()->user->returnUrl);
		}
		// display the login form
		$this->render('login',array('model'=>$model));
	}

	/**
	 * Logs out the current user and redirect to homepage.
	 */
	public function actionLogout()
	{
		Yii::app()->user->logout();
		$this->redirect(Yii::app()->homeUrl);
	}
}
