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

		 $objData = array();
		
		 $persone = Persons::model(); 
		
		 //COORDINATORI
		 $criteriaC = new CDbCriteria();
		 $criteriaC->with = array('personsMasters.master');
		 $criteriaC->together=true;
		 $criteriaC->condition='RoleID=18 AND master.Enabled=1';		 
		 $coordinatori = $persone->findAll($criteriaC);


		
        

		 foreach ($coordinatori as $coordinatore) {
		 
        	$arrayMastersCoordinatore = array();
        	
        	foreach ($coordinatore->personsMasters as $master) {
        		$arrayMastersCoordinatore[]=$master->MasterID;        		
        	}

        	 //PMA
	 		 $criteriaPma = new CDbCriteria();
	 		 $criteriaPma->with = array('personsMasters.master');
			 $criteriaPma->together=true;
			 $criteriaPma->addCondition('RoleID=11');
			 $criteriaPma->addInCondition('personsMasters.MasterID',$arrayMastersCoordinatore);

			 $listapma = $persone->findAll($criteriaPma);
			
			 //ANCONA
			 $criteriaPms = new CDbCriteria();
	 		 $criteriaPms->with = array('personsMasters.master','personsCities.city');
			 $criteriaPms->together=true;
			 $criteriaPms->addCondition('RoleID=15 AND personsCities.CityID=10');
			 $criteriaPms->addInCondition('personsMasters.MasterID',$arrayMastersCoordinatore);

			 $listapms = $persone->findAll($criteriaPms);

			 //OUTSIDER
			 $criteriaPmsOutsider = new CDbCriteria();
	 		 $criteriaPmsOutsider->with = array('personsMasters.master','personsCities');
			 $criteriaPmsOutsider->together=true;
			 $criteriaPmsOutsider->addCondition('RoleID=15 AND personsCities.CityID!=10');
			 $criteriaPmsOutsider->addInCondition('personsMasters.MasterID',$arrayMastersCoordinatore);

			 $listapmsOutsider = $persone->findAll($criteriaPmsOutsider);

			 //COSTRUISCO l' OBJECTDATA

			$objData[$coordinatore->PersonID] = array($coordinatore,$listapma,$listapms,$listapmsOutsider);

			//DISEGNO L^organigramma4

        	
		 }
		
		$this->draw($objData);
	
	}


	public function draw($objData){

		$numeroPma = 0;
		$posXCoordinatore = 0;
		$posYCoordinatore = 3;
		$spazioY = 2;

		$cellWidth = 5;



		$styleCoordinatore = array(
				'borders'=>array(
					'outline'=>array(
						'style'=>PHPExcel_Style_Border::BORDER_THIN
						),
					),
				'fill'=>array(
					'type'=>PHPExcel_Style_Fill::FILL_SOLID,
					'startcolor'=>array(
						'argb'=>'FFFF00'),
					),
				);

		$stylePma = array(
				'borders'=>array(
					'outline'=>array(
						'style'=>PHPExcel_Style_Border::BORDER_THIN
						),
					),
				'fill'=>array(
					'type'=>PHPExcel_Style_Fill::FILL_SOLID,
					'startcolor'=>array(
						'argb'=>'E0FFFF'),
					),
				);

		$stylePms = array(
				'borders'=>array(
					'outline'=>array(
						'style'=>PHPExcel_Style_Border::BORDER_THIN
						),
					),
				'fill'=>array(
					'type'=>PHPExcel_Style_Fill::FILL_SOLID,
					'startcolor'=>array(
						'argb'=>'E0FFBF'),
					),
				);


		$objPHPExcel = new PHPExcel();
       
        $sheet = $objPHPExcel->getActiveSheet()->setTitle('Organigramma SIDA');

       $acc = 0;

       foreach ($objData as $key => $value) {
       	    
       		$numeroPma = count($value[1]);
       		$posXCoordinatore = $posXCoordinatore + floor($numeroPma/2) + 1 + $acc;
       		$acc = floor($numeroPma/2) + 1 ;

       		$sheet->setCellValueByColumnAndRow($posXCoordinatore, $posYCoordinatore,  "COORDINATORE\n" . $value[0]->FirstName . " " . $value[0]->LastName);

       		$el =  $sheet->getColumnDimensionByColumn($posXCoordinatore);
       		$el->setAutoSIze(true);

       		$el = $sheet->getStyleByColumnAndRow($posXCoordinatore, $posYCoordinatore);
       		$el->applyFromArray($styleCoordinatore);
       		$el->getAlignment()->setWrapText(true);
        	$el->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        	$el->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        




       		$posXpma = $posXCoordinatore - ($acc - 1);
       		$posYpma = $posYCoordinatore + $spazioY;
       		//STAMPIAMO I PMA
       		foreach ($value[1] as $pma) {
       			$listaMasters = "";
       			foreach ($pma->personsMasters as $master) {

       				$listaMasters .= $master->master->ShortDescription . "\n";
       			}


       			$sheet->setCellValueByColumnAndRow( $posXpma , $posYpma ,  "PMA\n");
       			$sheet->setCellValueByColumnAndRow($posXpma , $posYpma + 1,$pma->FirstName . " " . $pma->LastName . "\n" );
       			$sheet->setCellValueByColumnAndRow($posXpma , $posYpma + 2, $listaMasters);

       			$elpma =  $sheet->getColumnDimensionByColumn($posXpma);
       			$elpma->setAutoSIze(true);

       			
       			$elpma = $sheet->getStyleByColumnAndRow($posXpma, $posYpma);
       			$elpma->applyFromArray($stylePma);
       			$elpma->getAlignment()->setWrapText(true);
        		$elpma->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        		$elpma->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);

        		$elpma = $sheet->getStyleByColumnAndRow($posXpma, $posYpma + 1);
       			$elpma->applyFromArray($stylePma);
       			$elpma->getAlignment()->setWrapText(true);
        		$elpma->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        		$elpma->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);

        		$elpma = $sheet->getStyleByColumnAndRow($posXpma, $posYpma + 2);
       			$elpma->applyFromArray($stylePma);
       			$elpma->getAlignment()->setWrapText(true);
        		$elpma->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        		$elpma->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);
       			
        		
        		//

        		foreach ($pma->personsMasters as $masterpma) {
        			

       				foreach ($value[2] as $pmsAncona2) {

       					
       					
       					foreach ($pmsAncona2->personsMasters as $masterpms) {
       						if($masterpms->MasterID == $masterpma->MasterID){
       							$sheet->setCellValueByColumnAndRow( $posXpma , $posYpma + 4 ,  "PMS ANCONA\n");
       							$sheet->setCellValueByColumnAndRow( $posXpma , $posYpma + 5,$pmsAncona2->FirstName . " " . $pmsAncona2->LastName);

       								 	$elpms =  $sheet->getColumnDimensionByColumn($posXpma);
						       			$elpms->setAutoSIze(true);

						       			
						       			$elpms = $sheet->getStyleByColumnAndRow($posXpma , $posYpma + 4);
						       			$elpms->applyFromArray($stylePms);
						       			$elpms->getAlignment()->setWrapText(true);
						        		$elpms->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
						        		$elpms->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);

						        		$elpms = $sheet->getStyleByColumnAndRow($posXpma , $posYpma + 5);
						       			$elpms->applyFromArray($stylePms);
						       			$elpms->getAlignment()->setWrapText(true);
						        		$elpms->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
						        		$elpms->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);




       							break 2;
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
       			$sheet->setCellValueByColumnAndRow( $posXpmsOutsider , $posYpmsOutsider ,  "PMS " . $pmsOutsider->personsCities[0]->city->Name . "\n");
       			$sheet->setCellValueByColumnAndRow($posXpmsOutsider , $posYpmsOutsider + 1,$pmsOutsider->FirstName . " " . $pmsOutsider->LastName );

       			$elpmsoutsider =  $sheet->getColumnDimensionByColumn($posXpmsOutsider);
       			$elpmsoutsider->setAutoSIze(true);

       			
       			$elpmsoutsider = $sheet->getStyleByColumnAndRow($posXpmsOutsider, $posYpmsOutsider);
       			$elpmsoutsider->applyFromArray($stylePms);
       			$elpmsoutsider->getAlignment()->setWrapText(true);
        		$elpmsoutsider->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        		$elpmsoutsider->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);

        		$elpmsoutsider = $sheet->getStyleByColumnAndRow($posXpmsOutsider, $posYpmsOutsider + 1);
       			$elpmsoutsider->applyFromArray($stylePms);
       			$elpmsoutsider->getAlignment()->setWrapText(true);
        		$elpmsoutsider->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        		$elpmsoutsider->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);
       			
       				
       			$posYpmsOutsider +=3;
       		}




       }


        $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
        $objWriter->save(Yii::app()->basePath . '/../files/exports/export.xlsx');

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
